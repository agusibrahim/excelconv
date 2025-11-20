package main

import (
	"encoding/json"
	"fmt"
	"io"
	"log"
	"math"
	"net/http"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	// Library wrapper libxls
	"github.com/godzie44/go-xls/xls"
	"github.com/xuri/excelize/v2"
)

const port = ":3020"

type ColumnMapping struct {
	Key      string
	Aliases  []string
	Required bool
}

var columnMappings = []ColumnMapping{
	// Tambahkan "no.pol" di sini
	{Key: "nopol", Aliases: []string{"licenseplate", "nopolisi", "nopol", "plate", "vehicleplate", "no.pol", "no.p"}, Required: true},

	{Key: "mobil", Aliases: []string{"unit", "assettype", "merk", "type", "jeniskendaraan", "mobil", "jenis", "typeunit", "jeniskendaraan", "vehicle"}},
	{Key: "lesing", Aliases: []string{"lesing", "leasing", "lesng", "finance", "financing"}},

	// Tambahkan "dd" di sini jika itu artinya overdue
	{Key: "ovd", Aliases: []string{"overdue", "ovd", "daysoverdue", "overdu", "hari", "keterlambatan", "dayslate", "dd"}},

	{Key: "saldo", Aliases: []string{"saldo", "credit", "balance", "amount", "remaining"}},
	{Key: "cabang", Aliases: []string{"branchfullname", "cabang", "branch", "office", "location"}},
	{Key: "nama", Aliases: []string{"ket", "keterangan", "catatan", "cat", "nama"}}, // Tambahkan "nama" juga
	{Key: "noka", Aliases: []string{"chasisno", "nomorrangka", "norangka", "no.rangka", "noka", "chassis", "frame"}},
	{Key: "nosin", Aliases: []string{"nomesin", "nomormesin", "no.mesin", "nosin", "engine", "engineno"}},
}

var outputOrder = []string{"nopol", "mobil", "lesing", "ovd", "saldo", "cabang", "nama", "noka", "nosin"}

func main() {
	if _, err := os.Stat("uploads"); os.IsNotExist(err) {
		os.Mkdir("uploads", 0755)
	}

	http.HandleFunc("/upload", uploadHandler)

	fmt.Printf("Server running on port %s\n", port)
	if err := http.ListenAndServe(port, nil); err != nil {
		log.Fatal(err)
	}
}

func uploadHandler(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
		return
	}

	r.ParseMultipartForm(20 << 20)

	file, handler, err := r.FormFile("excelFile")
	if err != nil {
		http.Error(w, "No file uploaded", http.StatusBadRequest)
		return
	}
	defer file.Close()

	ext := strings.ToLower(filepath.Ext(handler.Filename))
	if ext != ".xlsx" && ext != ".xls" {
		http.Error(w, "Invalid file type. Only .xlsx and .xls are allowed", http.StatusBadRequest)
		return
	}

	filename := filepath.Join("uploads", fmt.Sprintf("temp-%s", handler.Filename))
	dst, err := os.Create(filename)
	if err != nil {
		http.Error(w, "Unable to save file", http.StatusInternalServerError)
		return
	}

	if _, err := io.Copy(dst, file); err != nil {
		dst.Close()
		http.Error(w, "Unable to write file", http.StatusInternalServerError)
		return
	}
	dst.Close()

	defer func() {
		os.Remove(filename)
	}()

	var rawRows [][][]string
	var readErr error

	if ext == ".xlsx" {
		rawRows, readErr = readXLSX(filename)
	} else {
		// Panggil implementasi libxls
		rawRows, readErr = readXLS(filename)
		log.Println("Read XLS:", readErr)
		log.Println("Read XLS:", rawRows)
	}

	if readErr != nil {
		// Log error detail ke console untuk debugging
		log.Printf("Read Error: %v", readErr)
		http.Error(w, fmt.Sprintf("Error reading file: %v", readErr), http.StatusInternalServerError)
		return
	}

	data := processRawRows(rawRows)

	if len(data) == 0 {
		http.Error(w, `{"error": "No valid data found in the file"}`, http.StatusBadRequest)
		return
	}

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(data)
}

// ---------------------------------------------------------
// READERS
// ---------------------------------------------------------

func readXLSX(filePath string) ([][][]string, error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	var allSheetsRows [][][]string
	for _, sheetName := range f.GetSheetList() {
		rows, err := f.GetRows(sheetName)
		if err != nil {
			continue
		}
		allSheetsRows = append(allSheetsRows, rows)
	}
	return allSheetsRows, nil
}

// readXLS menggunakan github.com/godzie44/go-xls (Binding C libxls)
// readXLS menggunakan github.com/godzie44/go-xls (Binding C libxls)
// Versi VERBOSE untuk debugging
func readXLS(filePath string) ([][][]string, error) {
	log.Printf("[readXLS] Memulai proses baca file: %s", filePath)

	// Buka workbook
	wb, err := xls.OpenFile(filePath, "UTF-8")
	if err != nil {
		log.Printf("[readXLS] CRITICAL: Gagal membuka file xls: %v", err)
		return nil, err
	}
	defer wb.Close()

	var allSheetsRows [][][]string

	// Iterate melalui sheet.
	// Kita batasi 20 sheet untuk safety, biasanya file excel tidak sebanyak itu
	for i := 0; i < 20; i++ {
		log.Printf("[readXLS] Mencoba membuka Sheet index [%d]...", i)

		sheet, err := wb.OpenWorkSheet(i)
		if err != nil {
			// Error di sini biasanya normal (artinya sheet sudah habis)
			log.Printf("[readXLS] Berhenti di Sheet [%d] (End of sheets atau error): %v", i, err)
			break
		}

		// Gunakan anonymous function untuk memastikan sheet.Close()
		// dieksekusi segera setelah sheet selesai diproses, bukan di akhir fungsi utama.
		func() {
			defer sheet.Close()

			var sheetRows [][]string
			log.Printf("[readXLS] Sheet [%d] terbuka. Membaca baris...", i)

			// Loop Rows
			for rIdx, row := range sheet.Rows {
				var rowData []string

				// Loop Cells
				for _, cell := range row.Cells {
					// Value.String() otomatis menangani tipe data basic
					val := cell.Value.String()
					rowData = append(rowData, val)
				}

				// LOGGING TAMBAHAN: Tampilkan baris pertama (Header) untuk debug mapping
				if rIdx == 0 {
					log.Printf("[readXLS] Sheet [%d] - Header Baris 0: %v", i, rowData)
				}

				// Tambahkan logic untuk skip baris kosong total jika perlu
				// if len(rowData) > 0 { ... }

				sheetRows = append(sheetRows, rowData)
			}

			log.Printf("[readXLS] Sheet [%d] selesai. Total Baris: %d", i, len(sheetRows))
			allSheetsRows = append(allSheetsRows, sheetRows)
		}()
	}

	log.Printf("[readXLS] Selesai membaca file. Total Sheet yang valid: %d", len(allSheetsRows))
	return allSheetsRows, nil
}

// ---------------------------------------------------------
// LOGIC PROCESSOR (Sama seperti sebelumnya)
// ---------------------------------------------------------
func processRawRows(allSheetsRows [][][]string) [][]string {
	var result [][]string

	for _, rows := range allSheetsRows {
		if len(rows) < 5 {
			continue
		}

		headerMap := make(map[string]int)
		headerRowIdx := -1

		limit := 10
		if len(rows) < limit {
			limit = len(rows)
		}

		for i := 0; i < limit; i++ {
			row := rows[i]
			for colIdx, cell := range row {
				if cell == "" {
					continue
				}
				cleanedCell := strings.ToLower(strings.ReplaceAll(cell, " ", ""))

				for _, mapping := range columnMappings {
					for _, alias := range mapping.Aliases {
						if strings.Contains(cleanedCell, alias) {
							if _, exists := headerMap[mapping.Key]; !exists {
								headerMap[mapping.Key] = colIdx
								headerRowIdx = i
							}
						}
					}
				}
			}
			if len(headerMap) > 0 {
				break
			}
		}

		if headerRowIdx == -1 {
			continue
		}

		for i := headerRowIdx + 1; i < len(rows); i++ {
			row := rows[i]
			record := make(map[string]string)

			for key, colIdx := range headerMap {
				var value string
				if colIdx < len(row) {
					value = row[colIdx]
				}

				if value != "" {
					val := strings.ReplaceAll(strings.TrimSpace(value), ",", ".")

					switch key {
					case "saldo":
						if num, err := strconv.ParseFloat(val, 64); err == nil {
							record[key] = fmt.Sprintf("%.0f", math.Round(num))
						} else {
							record[key] = val
						}
					case "nopol":
						record[key] = strings.ReplaceAll(val, " ", "")
					default:
						record[key] = val
					}
				}
			}

			if val, ok := record["nopol"]; ok && val != "" {
				var rowData []string
				for _, outKey := range outputOrder {
					if v, exists := record[outKey]; exists {
						rowData = append(rowData, v)
					} else {
						rowData = append(rowData, "")
					}
				}
				result = append(result, rowData)
			}
		}
	}
	return result
}
