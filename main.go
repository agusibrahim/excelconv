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

	"github.com/xuri/excelize/v2"
)

// Konfigurasi Port
const port = ":3020"

// Struct untuk mapping
type ColumnMapping struct {
	Key      string
	Aliases  []string
	Required bool
}

// Konfigurasi mapping kolom (sama seperti JS)
var columnMappings = []ColumnMapping{
	{Key: "nopol", Aliases: []string{"licenseplate", "nopolisi", "nopol", "plate", "vehicleplate"}, Required: true},
	{Key: "mobil", Aliases: []string{"unit", "assettype", "merk", "type", "jeniskendaraan", "mobil", "jenis", "typeunit", "jeniskendaraan", "vehicle"}},
	{Key: "lesing", Aliases: []string{"lesing", "leasing", "lesng", "finance", "financing"}},
	{Key: "ovd", Aliases: []string{"overdue", "ovd", "daysoverdue", "overdu", "hari", "keterlambatan", "dayslate"}},
	{Key: "saldo", Aliases: []string{"saldo", "credit", "balance", "amount", "remaining"}},
	{Key: "cabang", Aliases: []string{"branchfullname", "cabang", "branch", "office", "location"}},
	{Key: "nama", Aliases: []string{"ket", "keterangan", "catatan", "cat"}},
	{Key: "noka", Aliases: []string{"chasisno", "nomorrangka", "norangka", "no.rangka", "noka", "chassis", "frame"}},
	{Key: "nosin", Aliases: []string{"nomesin", "nomormesin", "no.mesin", "nosin", "engine", "engineno"}},
}

// Urutan output yang diinginkan
var outputOrder = []string{"nopol", "mobil", "lesing", "ovd", "saldo", "cabang", "nama", "noka", "nosin"}

func main() {
	// Buat folder uploads jika belum ada
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
	// Hanya method POST
	if r.Method != http.MethodPost {
		http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
		return
	}

	// Parse multipart form (max 10MB)
	err := r.ParseMultipartForm(10 << 20)
	if err != nil {
		http.Error(w, "Error parsing form", http.StatusBadRequest)
		return
	}

	// Ambil file dari form key 'excelFile'
	file, handler, err := r.FormFile("excelFile")
	if err != nil {
		http.Error(w, "No file uploaded", http.StatusBadRequest)
		return
	}
	defer file.Close()

	// Simpan file sementara
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

	// Hapus file setelah proses selesai (defer berjalan di akhir fungsi)
	defer func() {
		os.Remove(filename)
	}()

	// Proses Excel
	data, err := processExcelFile(filename)
	if err != nil {
		http.Error(w, err.Error(), http.StatusInternalServerError)
		return
	}

	if len(data) == 0 {
		http.Error(w, `{"error": "No valid data found in the file"}`, http.StatusBadRequest)
		return
	}

	// Return JSON
	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(data)
}

func processExcelFile(filePath string) ([][]string, error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("error opening file: %v", err)
	}
	defer f.Close()

	var result [][]string

	// Loop semua sheet
	for _, sheetName := range f.GetSheetList() {
		rows, err := f.GetRows(sheetName)
		if err != nil {
			continue
		}

		if len(rows) < 5 {
			continue
		}

		headerMap := make(map[string]int)
		headerRowIdx := -1

		// Cari header dalam 10 baris pertama
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
				// Bersihkan cell untuk pencocokan (lowercase, remove space)
				cleanedCell := strings.ToLower(strings.ReplaceAll(cell, " ", ""))
				
				for _, mapping := range columnMappings {
					for _, alias := range mapping.Aliases {
						if strings.Contains(cleanedCell, alias) {
							// Jika belum ada mapping untuk key ini, set
							if _, exists := headerMap[mapping.Key]; !exists {
								headerMap[mapping.Key] = colIdx
								headerRowIdx = i
							}
						}
					}
				}
			}
			// Jika kita menemukan setidaknya satu header yang valid, anggap baris ini header
			if len(headerMap) > 0 {
				break
			}
		}

		if headerRowIdx == -1 {
			continue
		}

		// Proses data rows
		for i := headerRowIdx + 1; i < len(rows); i++ {
			row := rows[i]
			record := make(map[string]string)

			for key, colIdx := range headerMap {
				var value string
				// Pastikan index kolom ada di baris ini (mencegah index out of range)
				if colIdx < len(row) {
					value = row[colIdx]
				}

				if value != "" {
					// Cleaning umum: trim space, replace koma dengan titik
					val := strings.ReplaceAll(strings.TrimSpace(value), ",", ".")

					switch key {
					case "saldo":
						// Parse float, round, convert back to string
						if num, err := strconv.ParseFloat(val, 64); err == nil {
							record[key] = fmt.Sprintf("%.0f", math.Round(num))
						} else {
							record[key] = val
						}
					case "nopol":
						// Hapus spasi
						record[key] = strings.ReplaceAll(val, " ", "")
					default:
						record[key] = val
					}
				}
			}

			// Hanya masukkan jika 'nopol' ada
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

	return result, nil
}
