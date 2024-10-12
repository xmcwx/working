// sheet2report.go

package main

import (
	"bufio"
	"context"
	"crypto/rand"
	"encoding/hex"
	"encoding/json"
	"flag"
	"fmt"
	"log"
	"net/http"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"sync"
	"syscall"
	"time"

	"github.com/xuri/excelize/v2"
	"golang.org/x/crypto/nacl/secretbox"
	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	"golang.org/x/term"
	"google.golang.org/api/docs/v1"
	"google.golang.org/api/drive/v3"
	"google.golang.org/api/option"
)

type ReportData struct {
	UserInputs      map[string]string
	SpreadsheetData map[string]string
}

type Config struct {
	LogDir          string
	SpreadsheetFile string
}

var (
	logger     *log.Logger
	tokenStore = &TokenStore{
		tokens: make(map[string]*oauth2.Token),
		mutex:  &sync.RWMutex{},
	}
	config Config
)

type TokenStore struct {
	tokens map[string]*oauth2.Token
	mutex  *sync.RWMutex
}

func (ts *TokenStore) Get(key string) (*oauth2.Token, bool) {
	ts.mutex.RLock()
	defer ts.mutex.RUnlock()
	token, exists := ts.tokens[key]
	return token, exists
}

func (ts *TokenStore) Set(key string, token *oauth2.Token) {
	ts.mutex.Lock()
	defer ts.mutex.Unlock()
	ts.tokens[key] = token
}

func initLogger(logDir string) {
	err := os.MkdirAll(logDir, os.ModePerm)
	if err != nil {
		log.Fatalf("Failed to create logs directory: %v", err)
	}

	logFileName := filepath.Join(logDir, fmt.Sprintf("sheet2report_%s.log", time.Now().Format("2006-01-02_15-04-05")))
	logFile, err := os.OpenFile(logFileName, os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0666)
	if err != nil {
		log.Fatalf("Failed to open log file: %v", err)
	}

	logger = log.New(logFile, "", log.Ldate|log.Ltime|log.Lshortfile)
	logger.Println("Logging initialized")
}

func initTokenStore() {
	go func() {
		for {
			time.Sleep(5 * time.Minute) // Backup every 5 minutes
			if err := backupTokens(); err != nil {
				logger.Printf("Failed to backup tokens: %v", err)
			}
		}
	}()

	// Load tokens from backup on startup
	if err := loadTokensFromBackup(); err != nil {
		logger.Printf("Failed to load tokens from backup: %v", err)
	}
}

func backupTokens() error {
	tokenStore.mutex.RLock()
	defer tokenStore.mutex.RUnlock()

	encryptedTokens, err := encryptTokens(tokenStore.tokens)
	if err != nil {
		return fmt.Errorf("failed to encrypt tokens: %v", err)
	}

	return os.WriteFile("tokens.enc", encryptedTokens, 0600)
}

func loadTokensFromBackup() error {
	encryptedTokens, err := os.ReadFile("tokens.enc")
	if err != nil {
		if os.IsNotExist(err) {
			return nil
		}
		return fmt.Errorf("failed to read tokens file: %v", err)
	}

	decryptedTokens, err := decryptTokens(encryptedTokens)
	if err != nil {
		return fmt.Errorf("failed to decrypt tokens: %v", err)
	}

	tokenStore.mutex.Lock()
	defer tokenStore.mutex.Unlock()
	tokenStore.tokens = decryptedTokens

	return nil
}

func encryptTokens(tokens map[string]*oauth2.Token) ([]byte, error) {
	var nonce [24]byte
	if _, err := rand.Read(nonce[:]); err != nil {
		return nil, err
	}

	key := getEncryptionKey()

	tokensJSON, err := json.Marshal(tokens)
	if err != nil {
		return nil, err
	}

	return secretbox.Seal(nonce[:], tokensJSON, &nonce, &key), nil
}

func decryptTokens(encrypted []byte) (map[string]*oauth2.Token, error) {
	var nonce [24]byte
	copy(nonce[:], encrypted[:24])
	key := getEncryptionKey()

	decrypted, ok := secretbox.Open(nil, encrypted[24:], &nonce, &key)
	if !ok {
		return nil, fmt.Errorf("decryption failed")
	}

	var tokens map[string]*oauth2.Token
	if err := json.Unmarshal(decrypted, &tokens); err != nil {
		return nil, err
	}

	return tokens, nil
}

func getEncryptionKey() [32]byte {
	keyHex := os.Getenv("TOKEN_ENCRYPTION_KEY")
	var keyBytes []byte
	var err error

	if len(keyHex) == 64 {
		keyBytes, err = hex.DecodeString(keyHex)
		if err != nil {
			log.Fatalf("Invalid TOKEN_ENCRYPTION_KEY: %v", err)
		}
	} else {
		// Generate a random key
		keyBytes = make([]byte, 32)
		if _, err := rand.Read(keyBytes); err != nil {
			log.Fatalf("Failed to generate random key: %v", err)
		}
		keyHex = hex.EncodeToString(keyBytes)
		logger.Printf("Generated new encryption key: %s", keyHex)
		logger.Println("Consider setting this as TOKEN_ENCRYPTION_KEY environment variable for future runs.")
	}

	var fixedKey [32]byte
	copy(fixedKey[:], keyBytes)
	return fixedKey
}

func main() {
	config = Config{
		LogDir:          "logs",
		SpreadsheetFile: "master_output.xlsx",
	}

	// Define a custom flag set
	flagSet := flag.NewFlagSet("sheet2report", flag.ExitOnError)

	// Define flags
	flagSet.StringVar(&config.LogDir, "logdir", config.LogDir, "Directory for log files")
	flagSet.StringVar(&config.SpreadsheetFile, "spreadsheet", config.SpreadsheetFile, "Path to the master spreadsheet")

	// Define a help flag
	help := flagSet.Bool("help", false, "Display help message")

	// Parse the flags
	flagSet.Parse(os.Args[1:])

	// Check if help flag is set
	if *help {
		displayHelp()
		return
	}

	initLogger(config.LogDir)
	initTokenStore()
	defer logger.Println("Program execution completed")

	clientID, clientSecret := getClientCredentials()
	templateID := getTemplateID()

	logger.Println("Loading master spreadsheet")
	masterSpreadsheet, err := excelize.OpenFile(config.SpreadsheetFile)
	if err != nil {
		logger.Fatalf("Failed to open master spreadsheet: %v", err)
	}

	sheets := masterSpreadsheet.GetSheetMap()
	if len(sheets) == 0 {
		logger.Fatal("The master spreadsheet contains no sheets")
	}

	logger.Println("Authenticating with Google Docs, Drive, and Script services")
	ctx := context.Background()

	// Updated: Now returning the authenticated client
	docsService, driveService, client, err := getServices(ctx, clientID, clientSecret)
	if err != nil {
		logger.Fatalf("Failed to create services: %v", err)
	}

	logger.Println("Fetching template placeholders")
	userInputPlaceholders, sheetPlaceholders, err := getTemplatePlaceholders(ctx, docsService, templateID)
	if err != nil {
		logger.Fatalf("Failed to get template placeholders: %v", err)
	}

	logger.Printf("User input placeholders: %v", userInputPlaceholders)
	logger.Printf("Sheet placeholders: %v", sheetPlaceholders)

	logger.Println("Getting user inputs")
	userInputs := getUserInputs(userInputPlaceholders)

	totalSheets := len(sheets)
	i := 0
	for _, sheetName := range sheets {
		i++
		fmt.Printf("Processing sheet %d of %d: %s\n", i, totalSheets, sheetName)
		// Updated: Passing the authenticated client to processSheet
		err := processSheet(ctx, docsService, driveService, client, masterSpreadsheet, sheetName, templateID, userInputs, sheetPlaceholders)
		if err != nil {
			logger.Printf("Failed to process sheet %s: %v", sheetName, err)
			fmt.Printf("Failed to process sheet %s: %v\n", sheetName, err)
		}
	}

	fmt.Println("All sheets processed. Check the log file for details.")
	fmt.Println("Report links have been saved to report_links.txt")
}

func findSectionIndices(doc *docs.Document, sectionTitle string) (int64, int64, error) {
	var startIndex, endIndex int64
	found := false

	for _, element := range doc.Body.Content {
		if element.Paragraph != nil {
			for _, paraElement := range element.Paragraph.Elements {
				if paraElement.TextRun != nil && strings.Contains(paraElement.TextRun.Content, sectionTitle) {
					startIndex = paraElement.StartIndex
					found = true
				}
			}
		}

		if found && element.Paragraph != nil {
			for _, paraElement := range element.Paragraph.Elements {
				if paraElement.TextRun != nil && paraElement.StartIndex > startIndex {
					endIndex = paraElement.StartIndex
					return startIndex, endIndex, nil
				}
			}
		}
	}

	if found {
		// If no end index is found, use the end of the document
		endIndex = doc.Body.Content[len(doc.Body.Content)-1].EndIndex
		return startIndex, endIndex, nil
	}

	return 0, 0, fmt.Errorf("section %s not found", sectionTitle)
}

// Function to clean up the vulnerability summary
func cleanVulnerabilitySummary(summary string) string {
	lines := strings.Split(summary, "\n")
	var cleanedLines []string

	for _, line := range lines {
		trimmedLine := strings.TrimSpace(line)
		if trimmedLine != "" {
			cleanedLines = append(cleanedLines, trimmedLine)
		}
	}

	return strings.Join(cleanedLines, "\n")
}

func processSheet(ctx context.Context, docsService *docs.Service, driveService *drive.Service, client *http.Client,
	masterSpreadsheet *excelize.File, sheetName, templateID string,
	userInputs map[string]string, sheetPlaceholders []string) error {

	// Get rows and headers from the sheet
	rows, err := masterSpreadsheet.GetRows(sheetName)
	if err != nil {
		return fmt.Errorf("failed to get rows from sheet %s: %v", sheetName, err)
	}

	if len(rows) < 2 {
		return fmt.Errorf("sheet %s has insufficient data (less than 2 rows)", sheetName)
	}

	headers := rows[0]
	reportData := ReportData{
		UserInputs:      userInputs,
		SpreadsheetData: make(map[string]string),
	}

	// Variables to track vulnerabilities
	hasEndOfLifeWindows := false
	hasVulnerableLinux := false

	// Identify indexes of relevant columns
	osIndex := -1
	sysdescrIndex := -1

	for i, header := range headers {
		switch strings.ToLower(header) {
		case "os":
			osIndex = i
		case "sysdescr":
			sysdescrIndex = i
		}
	}

	// Check if the necessary columns are present
	if osIndex == -1 {
		return fmt.Errorf("'os' column not found in sheet %s", sheetName)
	}
	if sysdescrIndex == -1 {
		return fmt.Errorf("'sysdescr' column not found in sheet %s", sheetName)
	}

	// Iterate over data rows
	for _, row := range rows[1:] {
		var osValue, sysdescrValue string
		if osIndex < len(row) {
			osValue = row[osIndex]
		}
		if sysdescrIndex < len(row) {
			sysdescrValue = row[sysdescrIndex]
		}

		// Trim whitespace and convert to lower case for comparison
		osValueTrim := strings.TrimSpace(strings.ToLower(osValue))
		sysdescrValueTrim := strings.TrimSpace(strings.ToLower(sysdescrValue))

		// Check for end-of-life Windows systems
		if strings.HasPrefix(osValueTrim, "microsoft") || strings.HasPrefix(osValueTrim, "windows") {
			if isWindowsEndOfLife(osValueTrim) {
				hasEndOfLifeWindows = true
			}
		}

		// Check for vulnerable Linux systems
		if strings.HasPrefix(sysdescrValueTrim, "linux") {
			// Since all Linux systems are considered vulnerable
			hasVulnerableLinux = true
		}
	}

	var vulnerabilityHeader string
	var vulnerabilitySummary string

	switch {
	case hasEndOfLifeWindows && hasVulnerableLinux:
		vulnerabilityHeader = "Address End of Life and Vulnerable Operating Systems and Kernels"
		vulnerabilitySummary = `
Ensure that all Windows and Linux endpoints are upgraded to the most current OS or Kernel versions to prevent exploitation of known vulnerabilities.

Hosts that are end of life (EOL) should be prioritized for upgrading or replacement as it is possible they are no longer patched against newer vulnerabilities.

If upgrading or replacing EOL hosts are not possible then isolate them from the rest of the network using network segmentation and ACLs, deploy Endpoint Detection and Response (EDR) where possible, and continuously monitor for signs of unauthorized activity.
`
	case hasEndOfLifeWindows:
		vulnerabilityHeader = "Address End of Life Windows Systems"
		vulnerabilitySummary = `
Ensure that all Windows endpoints are upgraded to the most current OS versions to prevent exploitation of known vulnerabilities.

Hosts that are end of life (EOL) should be prioritized for upgrading or replacement as it is possible they are no longer patched against newer vulnerabilities.

If upgrading or replacing EOL hosts are not possible then isolate them from the rest of the network using network segmentation and ACLs, deploy Endpoint Detection and Response (EDR) where possible, and continuously monitor for signs of unauthorized activity.
`
	case hasVulnerableLinux:
		vulnerabilityHeader = "Address Vulnerable Linux Kernels"
		vulnerabilitySummary = `
Ensure that all Linux endpoints are upgraded to the most current Kernel versions to prevent exploitation of known vulnerabilities.

If upgrading or replacing vulnerable hosts are not possible then isolate them from the rest of the network using network segmentation and ACLs, deploy Endpoint Detection and Response (EDR) where possible, and continuously monitor for signs of unauthorized activity.
`
	default:
		vulnerabilityHeader = "No Vulnerabilities Detected"
		vulnerabilitySummary = `
The client has neither end-of-life Windows systems nor vulnerable Linux kernels.
`
	}

	// Clean the summary
	cleanedSummary := cleanVulnerabilitySummary(vulnerabilitySummary)

	reportData.SpreadsheetData["VulnerabilityHeader"] = vulnerabilityHeader
	reportData.SpreadsheetData["VulnerabilitySummary"] = cleanedSummary

	// Find the index of the "Unique Device Count" column
	uniqueDeviceCountIndex := -1
	for i, header := range headers {
		if strings.EqualFold(header, "unique device count") {
			uniqueDeviceCountIndex = i
			break
		}
	}

	// Extract the value from the first data row under the "Unique Device Count" column
	if uniqueDeviceCountIndex != -1 {
		if uniqueDeviceCountIndex < len(rows[1]) {
			uniqueDeviceCount := rows[1][uniqueDeviceCountIndex]
			reportData.SpreadsheetData["UniqueDeviceCount"] = uniqueDeviceCount
		} else {
			reportData.SpreadsheetData["UniqueDeviceCount"] = "0"
			logger.Printf("No data found under 'Unique Device Count' column in sheet %s", sheetName)
		}
	} else {
		reportData.SpreadsheetData["UniqueDeviceCount"] = "N/A"
		logger.Printf("'Unique Device Count' column not found in sheet %s", sheetName)
	}

	// Process other placeholders
	for _, placeholder := range sheetPlaceholders {
		if placeholder == "VulnerabilityHeader" || placeholder == "VulnerabilitySummary" || placeholder == "UniqueDeviceCount" {
			continue // We've already handled these
		}
		found := false
		for j, header := range headers {
			if strings.EqualFold(header, placeholder) && j < len(rows[1]) {
				reportData.SpreadsheetData[placeholder] = rows[1][j]
				found = true
				break
			}
		}
		if !found {
			logger.Printf("Warning: Placeholder %s not found in sheet %s", placeholder, sheetName)
			reportData.SpreadsheetData[placeholder] = "N/A"
		}
	}

	logger.Printf("Generating report for sheet: %s", sheetName)
	docID, err := generateGoogleDocsReport(ctx, docsService, driveService, templateID, reportData)
	if err != nil {
		return fmt.Errorf("failed to generate report for sheet %s: %v", sheetName, err)
	}

	clientName := reportData.SpreadsheetData["Client Name"]
	if clientName == "" {
		clientName = "Unnamed Client"
	}

	err = writeLinksToFile(clientName, docID)
	if err != nil {
		logger.Printf("Warning: Failed to write link to file: %v", err)
	}

	fmt.Printf("Report generated for sheet %s: https://docs.google.com/document/d/%s/edit\n", sheetName, docID)

	// Remove the TOC update functionality
	// Previously, there was a call to updateTOC function here

	return nil
}

func isWindowsEndOfLife(osValue string) bool {
	eolWindowsVersions := []string{
		"windows xp",
		"windows vista",
		"windows 7",
		"windows 8",
		"windows 8.1",
		"windows server 2003",
		"windows server 2008",
		"windows server 2008 r2",
		"windows server 2012",
		"windows server 2012 r2",
		// Add other EOL versions as needed
	}

	for _, eolVersion := range eolWindowsVersions {
		if strings.Contains(osValue, eolVersion) {
			return true
		}
	}
	return false
}

func getClientCredentials() (string, string) {
	clientID := os.Getenv("GOOGLE_CLIENT_ID")
	clientSecret := os.Getenv("GOOGLE_CLIENT_SECRET")

	if clientID == "" || clientSecret == "" {
		fmt.Println("Google OAuth credentials not found in environment variables.")
		fmt.Print("Enter Client ID: ")
		fmt.Scanln(&clientID)
		fmt.Print("Enter Client Secret: ")
		clientSecretBytes, _ := term.ReadPassword(int(syscall.Stdin))
		clientSecret = string(clientSecretBytes)
		fmt.Println()
	}

	return clientID, clientSecret
}

func getTemplateID() string {
	templateID := os.Getenv("GOOGLE_TEMPLATE_ID")
	if templateID == "" {
		fmt.Print("Enter Google Docs Template ID: ")
		fmt.Scanln(&templateID)
	}
	return templateID
}

func getUserID() string {
	// For simplicity, we're using a fixed user ID.
	return "default_user"
}

func getUserInputs(placeholders []string) map[string]string {
	inputs := make(map[string]string)

	reader := bufio.NewReader(os.Stdin)
	for _, placeholder := range placeholders {
		fmt.Printf("Enter value for %s: ", placeholder)
		answer, _ := reader.ReadString('\n')
		answer = strings.TrimSpace(answer)
		inputs[placeholder] = answer
		logger.Printf("User input for %s: %s", placeholder, answer)
	}

	return inputs
}

func getServices(ctx context.Context, clientID, clientSecret string) (*docs.Service, *drive.Service, *http.Client, error) {
	config := &oauth2.Config{
		ClientID:     clientID,
		ClientSecret: clientSecret,
		Endpoint:     google.Endpoint,
		RedirectURL:  "http://localhost:8080",
		Scopes: []string{
			docs.DocumentsScope,
			drive.DriveScope,
			// Removed: script.ScriptProjectsScope, // Removed scope for Apps Script API
		},
	}

	client := getClient(config)

	docsService, err := docs.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return nil, nil, nil, fmt.Errorf("unable to create Docs client: %v", err)
	}

	driveService, err := drive.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return nil, nil, nil, fmt.Errorf("unable to create Drive client: %v", err)
	}

	// Return the authenticated client
	return docsService, driveService, client, nil
}

func getClient(config *oauth2.Config) *http.Client {
	userID := getUserID()
	token, exists := tokenStore.Get(userID)
	if !exists || token == nil || !token.Valid() {
		token = getTokenFromWeb(config)
		tokenStore.Set(userID, token)
	}
	return config.Client(context.Background(), token)
}

func getTokenFromWeb(config *oauth2.Config) *oauth2.Token {
	// Generate a random state string for security
	state := generateStateString()

	// Create a channel to receive the authorization code
	codeCh := make(chan string)

	// Set up the redirect URI
	redirectURI := "localhost:8080"
	config.RedirectURL = "http://" + redirectURI

	// Start a local HTTP server
	srv := &http.Server{Addr: redirectURI}

	// Set up the handler
	http.HandleFunc("/", func(w http.ResponseWriter, r *http.Request) {
		// Log the received request
		logger.Println("Received OAuth callback request")

		// Validate the state parameter
		if r.URL.Query().Get("state") != state {
			http.Error(w, "State parameter doesn't match", http.StatusBadRequest)
			logger.Println("Invalid state parameter")
			return
		}
		code := r.URL.Query().Get("code")
		if code == "" {
			http.Error(w, "Code not found", http.StatusBadRequest)
			logger.Println("Authorization code not found in the request")
			return
		}

		// Send response to the browser
		fmt.Fprintln(w, "Authorization successful! You can close this window.")

		// Send the code to the channel
		codeCh <- code
	})

	// Start the server in a goroutine
	go func() {
		logger.Printf("Starting local server on %s", redirectURI)
		if err := srv.ListenAndServe(); err != nil && err != http.ErrServerClosed {
			logger.Fatalf("Failed to start server: %v", err)
		}
	}()

	// Generate the authorization URL
	authURL := config.AuthCodeURL(state, oauth2.AccessTypeOffline)
	fmt.Printf("Go to the following link in your browser:\n%v\n", authURL)
	logger.Printf("Authorization URL: %s", authURL)

	// Wait for the authorization code
	code := <-codeCh
	logger.Println("Authorization code received")

	// Shut down the server
	logger.Println("Shutting down the server")
	ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
	defer cancel()
	if err := srv.Shutdown(ctx); err != nil {
		logger.Fatalf("Server Shutdown Failed: %v", err)
	}

	// Exchange the authorization code for a token
	tok, err := config.Exchange(context.Background(), code)
	if err != nil {
		logger.Fatalf("Unable to exchange code for token: %v", err)
	}
	logger.Println("Access token obtained successfully")

	return tok
}

// generateStateString generates a random state string for OAuth 2.0 authorization
func generateStateString() string {
	b := make([]byte, 16)
	if _, err := rand.Read(b); err != nil {
		logger.Fatalf("Unable to generate state string: %v", err)
	}
	return fmt.Sprintf("%x", b)
}

func removeDuplicates(slice []string) []string {
	keys := make(map[string]bool)
	list := []string{}
	for _, entry := range slice {
		if _, value := keys[entry]; !value {
			keys[entry] = true
			list = append(list, entry)
		}
	}
	return list
}

func getTemplatePlaceholders(ctx context.Context, srv *docs.Service, templateID string) ([]string, []string, error) {
	doc, err := srv.Documents.Get(templateID).Do()
	if err != nil {
		return nil, nil, fmt.Errorf("unable to retrieve the template document: %v", err)
	}

	var userInputPlaceholders, sheetPlaceholders []string
	for _, element := range doc.Body.Content {
		userInput, sheet := extractPlaceholders(element)
		userInputPlaceholders = append(userInputPlaceholders, userInput...)
		sheetPlaceholders = append(sheetPlaceholders, sheet...)
	}

	// Remove duplicates
	userInputPlaceholders = removeDuplicates(userInputPlaceholders)
	sheetPlaceholders = removeDuplicates(sheetPlaceholders)

	// Ensure "VulnerabilityHeader", "VulnerabilitySummary", and "UniqueDeviceCount" are in sheetPlaceholders
	sheetPlaceholders = append(sheetPlaceholders, "VulnerabilityHeader", "VulnerabilitySummary", "UniqueDeviceCount")

	// Remove "VulnerabilityHeader", "VulnerabilitySummary", and "UniqueDeviceCount" from userInputPlaceholders if they exist
	userInputPlaceholders = removeItems(userInputPlaceholders, []string{"VulnerabilityHeader", "VulnerabilitySummary", "UniqueDeviceCount"})

	logger.Printf("Detected user input placeholders: %v", userInputPlaceholders)
	logger.Printf("Detected sheet placeholders: %v", sheetPlaceholders)

	return userInputPlaceholders, sheetPlaceholders, nil
}

func removeItems(slice []string, items []string) []string {
	result := make([]string, 0, len(slice))
	itemSet := make(map[string]struct{}, len(items))
	for _, item := range items {
		itemSet[item] = struct{}{}
	}
	for _, s := range slice {
		if _, exists := itemSet[s]; !exists {
			result = append(result, s)
		}
	}
	return result
}

func extractPlaceholders(element *docs.StructuralElement) ([]string, []string) {
	var userInputPlaceholders, sheetPlaceholders []string
	re := regexp.MustCompile(`{{(.+?)}}`)

	if element.Paragraph != nil {
		for _, paraElement := range element.Paragraph.Elements {
			if paraElement.TextRun != nil {
				matches := re.FindAllStringSubmatch(paraElement.TextRun.Content, -1)
				for _, match := range matches {
					if strings.HasPrefix(match[1], "SHEET:") {
						sheetPlaceholders = append(sheetPlaceholders, strings.TrimPrefix(match[1], "SHEET:"))
					} else {
						userInputPlaceholders = append(userInputPlaceholders, match[1])
					}
				}
			}
		}
	} else if element.Table != nil {
		for _, row := range element.Table.TableRows {
			for _, cell := range row.TableCells {
				for _, cellContent := range cell.Content {
					userInput, sheet := extractPlaceholders(cellContent)
					userInputPlaceholders = append(userInputPlaceholders, userInput...)
					sheetPlaceholders = append(sheetPlaceholders, sheet...)
				}
			}
		}
	} else if element.SectionBreak != nil {
		// Handle other structural elements if necessary
	}

	return userInputPlaceholders, sheetPlaceholders
}

func generateGoogleDocsReport(ctx context.Context, docsService *docs.Service, driveService *drive.Service,
	templateID string, data ReportData) (string, error) {
	logger.Println("Copying template document")

	// Get the client name from the spreadsheet data
	clientName := data.SpreadsheetData["Client Name"]
	if clientName == "" {
		clientName = "Unnamed Client" // Default name if Client Name is not found
	}

	// Copy the template document
	copiedFile, err := driveService.Files.Copy(templateID, &drive.File{
		Name: fmt.Sprintf("Information Risk Analysis - %s", clientName),
	}).Do()
	if err != nil {
		return "", fmt.Errorf("unable to copy template document: %v", err)
	}

	copiedDocID := copiedFile.Id
	logger.Printf("Copied document ID: %s", copiedDocID)

	logger.Println("Replacing placeholders in the copied document")
	requests := []*docs.Request{}

	// Replace user input placeholders
	for key, value := range data.UserInputs {
		requests = append(requests, &docs.Request{
			ReplaceAllText: &docs.ReplaceAllTextRequest{
				ContainsText: &docs.SubstringMatchCriteria{
					Text:      "{{" + key + "}}",
					MatchCase: true,
				},
				ReplaceText: value,
			},
		})
	}

	// Replace spreadsheet data placeholders
	for key, value := range data.SpreadsheetData {
		placeholder := "{{SHEET:" + key + "}}"
		if key == "VulnerabilityHeader" || key == "VulnerabilitySummary" || key == "UniqueDeviceCount" {
			// Since these placeholders don't have 'SHEET:', adjust accordingly
			placeholder = "{{" + key + "}}"
		}
		requests = append(requests, &docs.Request{
			ReplaceAllText: &docs.ReplaceAllTextRequest{
				ContainsText: &docs.SubstringMatchCriteria{
					Text:      placeholder,
					MatchCase: true,
				},
				ReplaceText: value,
			},
		})
	}

	// If no vulnerabilities are detected, remove the "Recommendations" section
	if data.SpreadsheetData["VulnerabilityHeader"] == "No Vulnerabilities Detected" {
		doc, err := docsService.Documents.Get(copiedDocID).Do()
		if err != nil {
			return "", fmt.Errorf("unable to retrieve the document: %v", err)
		}

		startIndex, endIndex, err := findSectionIndices(doc, "Recommendations")
		if err != nil {
			return "", fmt.Errorf("unable to find 'Recommendations' section: %v", err)
		}

		requests = append(requests, &docs.Request{
			DeleteContentRange: &docs.DeleteContentRangeRequest{
				Range: &docs.Range{
					StartIndex: startIndex,
					EndIndex:   endIndex,
				},
			},
		})
	}

	// Apply the requests
	_, err = docsService.Documents.BatchUpdate(copiedDocID, &docs.BatchUpdateDocumentRequest{
		Requests: requests,
	}).Do()
	if err != nil {
		return "", fmt.Errorf("unable to update document: %v", err)
	}

	logger.Printf("Report generated successfully: %s", copiedDocID)
	return copiedDocID, nil
}

func displayHelp() {
	helpMessage := `
Sheet2Report - Generate Google Docs reports from spreadsheet data

Usage: sheet2report [options]

Options:
  -h, --help                    Display this help message
  -logdir string                Directory for log files (default "logs")
  -spreadsheet string           Path to the master spreadsheet (default "master_output.xlsx")

Environment Variables:
  GOOGLE_CLIENT_ID              Your Google OAuth 2.0 Client ID
  GOOGLE_CLIENT_SECRET          Your Google OAuth 2.0 Client Secret
  GOOGLE_TEMPLATE_ID            Your Google Docs template ID
  TOKEN_ENCRYPTION_KEY          Optional: A 64-character hex string (32 bytes) for token encryption.
                                If not provided, a random key will be generated.

Description:
  This tool generates individual Google Docs reports based on data from a master spreadsheet
  with multiple sheets and a Google Docs template. It prompts the user for additional inputs
  and combines this information to create customized reports, one for each sheet.

Requirements:
  - master_output.xlsx: The Excel file containing multiple sheets with data for the reports.
  - A Google Docs template with placeholders for report content.
  - Enabled Google Docs API and Google Drive API in your Google Cloud project.
  - Appropriate OAuth 2.0 credentials with the necessary scopes.

Process:
  1. The program reads the master spreadsheet.
  2. It prompts for Google OAuth credentials and template ID if not provided via environment variables.
  3. It extracts placeholders from the template document.
  4. It prompts the user for input for non-sheet placeholders.
  5. It processes each sheet separately.
     For each sheet, it generates a report by copying the template
     and replacing placeholders with data from the spreadsheet and user inputs.
  6. The generated reports are saved as individual Google Docs.

Placeholders:
  - User input placeholders: {{Placeholder Name}}
  - Spreadsheet data placeholders: {{SHEET:Column Name}}
  - Vulnerability header placeholder: {{VulnerabilityHeader}} (dynamic content based on analysis)
  - Vulnerability summary placeholder: {{VulnerabilitySummary}} (dynamic content based on analysis)
  - Unique device count placeholder: {{UniqueDeviceCount}} (from spreadsheet data)

Scopes:
  - The application uses the following OAuth scopes:
    - https://www.googleapis.com/auth/documents
    - https://www.googleapis.com/auth/drive

Notes:
	- The application securely stores OAuth tokens locally, encrypting them for added security.
	- It generates a report for each sheet in the master spreadsheet.
	- The program logs its activities to a file in the specified log directory.
	- Report links are saved to a 'report_links.txt' file for easy access.
	- The tool automatically detects end-of-life Windows systems and vulnerable Linux kernels.
  
Example usage:
	sheet2report -logdir ./logs -spreadsheet ./data/master_sheet.xlsx
  
For more information or support, please refer to the documentation or contact the developer.`

	fmt.Println(helpMessage)
}

func writeLinksToFile(clientName, docID string) error {
	file, err := os.OpenFile("report_links.txt", os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0644)
	if err != nil {
		return fmt.Errorf("failed to open report_links.txt: %v", err)
	}
	defer file.Close()

	link := fmt.Sprintf("https://docs.google.com/document/d/%s/edit", docID)
	_, err = file.WriteString(fmt.Sprintf("Client: %s\nReport Link: %s\n\n", clientName, link))
	if err != nil {
		return fmt.Errorf("failed to write to report_links.txt: %v", err)
	}

	return nil
}
