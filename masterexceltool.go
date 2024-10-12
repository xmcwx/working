package main

import (
	"bufio"
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"time"

	"github.com/xuri/excelize/v2"
)

var (
	safeClients      []string
	safeClientsMutex sync.Mutex
	detailedLogger   *log.Logger
)

func main() {
	// Define command-line flags
	inputDir := flag.String("input", "", "Directory with input spreadsheets (required)")
	outputFile := flag.String("output", "master_output.xlsx", "Output master spreadsheet file path")
	targetSheet := flag.String("sheet", "Sheet1", "The specific sheet to extract data from (default: Sheet1)")
	flag.Parse()

	// Show help message if input directory is not provided
	if *inputDir == "" {
		flag.Usage()
		os.Exit(1)
	}

	// Check if input directory exists
	if _, err := os.Stat(*inputDir); os.IsNotExist(err) {
		log.Fatalf("Input directory does not exist: %s", *inputDir)
	}

	// Set up detailed logging
	detailedLogFile := strings.TrimSuffix(*outputFile, filepath.Ext(*outputFile)) + "_detailed.log"
	f, err := os.Create(detailedLogFile)
	if err != nil {
		log.Fatalf("Failed to create detailed log file: %v", err)
	}
	defer f.Close()
	detailedLogger = log.New(f, "", log.Ldate|log.Ltime)

	detailedLogger.Println("Script started")
	detailedLogger.Printf("Input directory: %s", *inputDir)
	detailedLogger.Printf("Output file: %s", *outputFile)
	detailedLogger.Printf("Target sheet: %s", *targetSheet)

	// Create a new Excel file for the master output
	masterFile := excelize.NewFile()

	// Process the first spreadsheet to get column selections
	var columnsToCopy []string

	files, err := filepath.Glob(filepath.Join(*inputDir, "*.xlsx"))
	if err != nil {
		log.Fatalf("Error finding Excel files: %v", err)
	}

	safeClients = []string{}

	for _, path := range files {
		processFileWithTimeout(path, masterFile, &columnsToCopy, *targetSheet)
	}

	writeSafeClientsLog(*outputFile)

	// Save the master Excel file
	if err := masterFile.SaveAs(*outputFile); err != nil {
		log.Fatalf("failed to save master file: %v", err)
	}

	detailedLogger.Printf("Master spreadsheet saved to: %s", *outputFile)
	detailedLogger.Printf("Safe clients list saved to: %s", strings.TrimSuffix(*outputFile, filepath.Ext(*outputFile))+"_safe_clients.txt")
	detailedLogger.Printf("Detailed log saved to: %s", detailedLogFile)

	detailedLogger.Println("Script finished")

	fmt.Println("Master spreadsheet saved successfully to", *outputFile)
	fmt.Println("Detailed log saved to", detailedLogFile)
}

func processFileWithTimeout(path string, masterFile *excelize.File, columnsToCopy *[]string, targetSheet string) {
	done := make(chan bool)
	go func() {
		detailedLogger.Printf("Processing file: %s", path)

		// Open the spreadsheet
		f, err := excelize.OpenFile(path)
		if err != nil {
			detailedLogger.Printf("failed to open file %s: %v\n", path, err)
			done <- true
			return
		}
		defer func() {
			if err := f.Close(); err != nil {
				detailedLogger.Printf("failed to close file %s: %v\n", path, err)
			}
		}()

		// Extract client name from "account" sheet, cell A2
		clientName := extractClientName(f, path)
		if clientName == "" {
			// Use the default clientName derived from the file name if extraction fails
			fileName := filepath.Base(path[:len(path)-len(filepath.Ext(path))])
			parts := strings.SplitN(fileName, "_", 2)
			if len(parts) > 1 {
				clientName = parts[1]
			} else {
				clientName = fileName
			}
		}
		detailedLogger.Printf("Extracted client name: %s", clientName)

		// Get rows from the specific sheet in the original file
		rows, err := f.GetRows(targetSheet)
		if err != nil {
			detailedLogger.Printf("failed to get rows from sheet %s in file %s: %v\n", targetSheet, path, err)
			done <- true
			return
		}
		detailedLogger.Printf("Found %d rows in sheet %s of file %s", len(rows), targetSheet, path)

		if len(rows) == 0 {
			detailedLogger.Printf("No data found in sheet %s of file %s", targetSheet, path)
			done <- true
			return
		}

		// Build a map of header names to their indices
		headerIndexMap := make(map[string]int)
		for i, header := range rows[0] {
			headerIndexMap[header] = i
		}

		// Prompt user for columns to copy only for the first file
		if len(*columnsToCopy) == 0 {
			*columnsToCopy = promptForColumns(rows[0])
			detailedLogger.Printf("Selected columns: %v", *columnsToCopy)
		}

		// Create a new sheet in the master file named after the client name
		sheetName := clientName
		// Truncate sheet name if it's longer than 31 characters
		if len(sheetName) > 31 {
			sheetName = sheetName[:31]
		}
		// Ensure the sheet name is unique
		sheetName = getUniqueSheetName(masterFile, sheetName)
		index, err := masterFile.NewSheet(sheetName)
		if err != nil {
			detailedLogger.Printf("Failed to create new sheet %s: %v", sheetName, err)
			done <- true
			return
		}
		masterFile.SetActiveSheet(index)

		detailedLogger.Printf("Created new sheet: %s", sheetName)

		// Define Windows end-of-life versions
		endOfLifeWindowsVersions := []string{
			"Windows XP",
			"Windows Vista",
			"Windows 7",
			"Windows 8",
			"Windows Server 2003",
			"Windows Server 2008",
			"Windows Server 2012",
		}

		// Find the index of the "sysdescr" column if it exists
		sysdescrIndex := -1
		for i, header := range *columnsToCopy {
			if header == "sysdescr" {
				sysdescrIndex = i
				break
			}
		}

		// Add headers
		masterFile.SetCellValue(sheetName, "A1", "Client Name")
		masterFile.SetCellValue(sheetName, "B1", "OS Information")

		rowCounter := 1
		relevantRowFound := false

		uniqueDevices := make(map[string]bool)
		deviceNameColumnIndex := -1

		// Find the "device name" column
		for i, header := range rows[0] {
			if strings.EqualFold(header, "device name") {
				deviceNameColumnIndex = i
				break
			}
		}

		if deviceNameColumnIndex == -1 {
			detailedLogger.Printf("'device name' column not found in %s", path)
			done <- true
			return
		}

		// Count unique devices for all rows (except header)
		for _, row := range rows[1:] {
			if deviceNameColumnIndex < len(row) {
				deviceName := strings.TrimSpace(row[deviceNameColumnIndex])
				if deviceName != "" {
					uniqueDevices[deviceName] = true
				}
			}
		}

		uniqueCount := len(uniqueDevices)

		// Copy data from selected columns in the target sheet
		for rowIndex, row := range rows {
			if rowIndex == 0 {
				// Copy headers, starting from column C
				for colIndex, header := range *columnsToCopy {
					cell, _ := excelize.CoordinatesToCellName(colIndex+3, 1) // +3 because we start from column C
					if err := masterFile.SetCellValue(sheetName, cell, header); err != nil {
						detailedLogger.Printf("failed to set header value in master file sheet %s: %v\n", sheetName, err)
					}
				}
			} else {
				// Check if the row contains end-of-life Windows versions or Linux in sysdescr
				isRelevant := false
				for colIndex, header := range *columnsToCopy {
					colNum, ok := headerIndexMap[header]
					if !ok {
						continue
					}
					if colNum < len(row) {
						cellValue := strings.TrimSpace(row[colNum])

						// Check for end-of-life Windows versions
						for _, version := range endOfLifeWindowsVersions {
							if strings.Contains(strings.ToLower(cellValue), strings.ToLower(version)) {
								isRelevant = true
								break
							}
						}

						// Check for Linux in sysdescr column
						if sysdescrIndex != -1 && colIndex == sysdescrIndex {
							words := strings.Fields(cellValue)
							if len(words) > 0 && strings.ToLower(words[0]) == "linux" {
								isRelevant = true
							}
						}

						if isRelevant {
							break
						}
					}
				}

				// Copy data only if relevant (end-of-life Windows or Linux)
				if isRelevant {
					relevantRowFound = true
					rowCounter++
					// Add client name to column A
					cellA, _ := excelize.CoordinatesToCellName(1, rowCounter)
					masterFile.SetCellValue(sheetName, cellA, clientName)

					// Add "Vulnerable or End of Life OS or Linux Kernel found" to column B
					cellB, _ := excelize.CoordinatesToCellName(2, rowCounter)
					masterFile.SetCellValue(sheetName, cellB, "Vulnerable or End of Life OS or Linux Kernel found")

					for colIndex, header := range *columnsToCopy {
						colNum, ok := headerIndexMap[header]
						if !ok {
							continue
						}
						if colNum < len(row) {
							cellValue := strings.TrimSpace(row[colNum])
							cell, _ := excelize.CoordinatesToCellName(colIndex+3, rowCounter) // +3 because we start from column C
							if err := masterFile.SetCellValue(sheetName, cell, cellValue); err != nil {
								detailedLogger.Printf("failed to set cell value in master file sheet %s: %v\n", sheetName, err)
							}
						}
					}
					logRelevantRow(clientName, row)
				}
			}
		}

		// If no relevant rows were found, add a single row with the message and unique device count
		if !relevantRowFound {
			masterFile.SetCellValue(sheetName, "A2", clientName)
			masterFile.SetCellValue(sheetName, "B2", "No Vulnerable or End of Life OS or Linux Kernel found")

			uniqueCountColumn, _ := excelize.CoordinatesToCellName(len(*columnsToCopy)+3, 2)
			masterFile.SetCellValue(sheetName, uniqueCountColumn, uniqueCount)

			safeClientsMutex.Lock()
			safeClients = append(safeClients, clientName)
			safeClientsMutex.Unlock()
		}

		// Add "Unique device count" column header
		uniqueCountColumnIndex := len(*columnsToCopy) + 3 // +3 because we start from column C
		uniqueCountColumn, _ := excelize.CoordinatesToCellName(uniqueCountColumnIndex, 1)
		masterFile.SetCellValue(sheetName, uniqueCountColumn, "unique device count")

		// Set the unique count value in the new column for all rows
		for i := 2; i <= rowCounter; i++ {
			cell, _ := excelize.CoordinatesToCellName(uniqueCountColumnIndex, i)
			if i == 2 {
				masterFile.SetCellValue(sheetName, cell, uniqueCount)
			} else {
				masterFile.SetCellValue(sheetName, cell, "")
			}
		}

		detailedLogger.Printf("Finished processing file: %s", path)
		if relevantRowFound {
			detailedLogger.Printf("Found vulnerable systems in file: %s", path)
		} else {
			detailedLogger.Printf("No vulnerable systems found in file: %s", path)
		}

		done <- true
	}()
	select {
	case <-done:
		log.Printf("Finished processing file: %s", path)
	case <-time.After(5 * time.Minute):
		log.Printf("Processing timed out for file: %s", path)
		detailedLogger.Printf("Processing timed out for file: %s", path)
	}
}

func promptForColumns(headers []string) []string {
	fmt.Println("Available columns:")
	for i, header := range headers {
		fmt.Printf("%d. %s\n", i+1, header)
	}

	fmt.Print("Enter the numbers of the columns you want to copy (separated by spaces): ")
	reader := bufio.NewReader(os.Stdin)
	input, _ := reader.ReadString('\n')
	input = strings.TrimSpace(input)

	selectedIndices := strings.Fields(input)
	selectedColumns := make([]string, 0)

	for _, index := range selectedIndices {
		i := 0
		fmt.Sscanf(index, "%d", &i)
		if i > 0 && i <= len(headers) {
			selectedColumns = append(selectedColumns, headers[i-1])
		}
	}

	return selectedColumns
}

func getUniqueSheetName(f *excelize.File, baseName string) string {
	// Remove invalid characters from the base name
	invalidChars := []string{":", "\\", "/", "?", "*", "[", "]"}
	for _, char := range invalidChars {
		baseName = strings.ReplaceAll(baseName, char, "_")
	}

	name := baseName
	counter := 1
	for {
		index, err := f.GetSheetIndex(name)
		if err != nil || index == -1 {
			// Sheet doesn't exist, so this name is unique
			return name
		}
		// Sheet exists, try next name
		counter++
		trimmedBase := baseName
		if len(baseName) > 28 {
			trimmedBase = baseName[:28]
		}
		name = fmt.Sprintf("%s_%d", trimmedBase, counter)
	}
}

func writeSafeClientsLog(outputFile string) {
	logFileName := strings.TrimSuffix(outputFile, filepath.Ext(outputFile)) + "_safe_clients.txt"
	f, err := os.Create(logFileName)
	if err != nil {
		detailedLogger.Printf("Failed to create safe clients log file: %v", err)
		return
	}
	defer f.Close()

	writer := bufio.NewWriter(f)
	writer.WriteString("Clients without vulnerable or end of life OS or Linux kernel:\n")
	for _, client := range safeClients {
		writer.WriteString(client + "\n")
	}
	writer.Flush()

	detailedLogger.Printf("Safe clients log written to: %s", logFileName)
	detailedLogger.Printf("Number of safe clients: %d", len(safeClients))
}

func logRelevantRow(clientName string, rowData []string) {
	detailedLogger.Printf("Relevant row found for client %s:", clientName)
	for i, value := range rowData {
		detailedLogger.Printf("  Column %d: %s", i+1, value)
	}
}

// New function to extract client name from the "account" sheet
func extractClientName(f *excelize.File, path string) string {
	accountSheetName := "Account"
	cellAddress := "A2"

	// Check if "account" sheet exists
	wbSheets := f.GetSheetList()
	sheetExists := false
	for _, sheet := range wbSheets {
		if sheet == accountSheetName {
			sheetExists = true
			break
		}
	}

	if !sheetExists {
		detailedLogger.Printf("Sheet %s not found in file %s", accountSheetName, path)
		return ""
	}

	// Get value from cell A2 in "account" sheet
	cellValue, err := f.GetCellValue(accountSheetName, cellAddress)
	if err != nil {
		detailedLogger.Printf("Failed to get cell %s value from sheet %s in file %s: %v", cellAddress, accountSheetName, path, err)
		return ""
	}

	// Parse the cell value to extract the client name
	// Expected format: "123456: ClientName Ltd."
	parts := strings.SplitN(cellValue, ": ", 2)
	if len(parts) == 2 {
		clientName := parts[1]
		detailedLogger.Printf("Extracted client name '%s' from cell %s in sheet %s", clientName, cellAddress, accountSheetName)
		return clientName
	}

	detailedLogger.Printf("Unexpected format in cell %s in sheet %s in file %s. Expected 'ID: ClientName'. Got '%s'", cellAddress, accountSheetName, path, cellValue)
	return ""
}
