package main

import (
	"fmt"
	"io"
	"log"
	"math"
	"net/http"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	// JSON yang lebih cepat
	"github.com/goccy/go-json"

	// Library Excel
	"github.com/godzie44/go-xls/xls"
	"github.com/xuri/excelize/v2"
)

const port = ":3020"

// Struct Config
type ColumnMapping struct {
	Key      string
	Aliases  []string
	Required bool
}

// Mapping (Sudah diupdate sesuai kasus Anda)
var columnMappings = []ColumnMapping{
	{Key: "nopol", Aliases: []string{"licenseplate", "nopolisi", "nopol", "plate", "vehicleplate", "no.pol", "no.p"}, Required: true},
	{Key: "mobil", Aliases: []string{"unit", "assettype", "merk", "type", "jeniskendaraan", "mobil", "jenis", "typeunit", "jeniskendaraan", "vehicle"}},
	{Key: "lesing", Aliases: []string{"lesing", "leasing", "lesng", "finance", "financing"}},
	{Key: "ovd", Aliases: []string{"overdue", "ovd", "daysoverdue", "overdu", "hari", "keterlambatan", "dayslate", "dd"}},
	{Key: "saldo", Aliases: []string{"saldo", "credit", "balance", "amount", "remaining"}},
	{Key: "cabang", Aliases: []string{"branchfullname", "cabang", "branch", "office", "location"}},
	{Key: "nama", Aliases: []string{"ket", "keterangan", "catatan", "cat", "nama"}},
	{Key: "noka", Aliases: []string{"chasisno", "nomorrangka", "norangka", "no.rangka", "noka", "chassis", "frame"}},
	{Key: "nosin", Aliases: []string{"nomesin", "nomormesin", "no.mesin", "nosin", "engine", "engineno"}},
}

var outputOrder = []string{"nopol", "mobil", "lesing", "ovd", "saldo", "cabang", "nama", "noka", "nosin"}

func main() {
	if _, err := os.Stat("uploads"); os.IsNotExist(err) {
		os.Mkdir("uploads", 0755)
	}

	// Gunakan http.HandlerFunc untuk routing sederhana
	http.HandleFunc("/upload", uploadHandler)

	fmt.Printf("Server optimized running on port %s\n", port)
	if err := http.ListenAndServe(port, nil); err != nil {
		log.Fatal(err)
	}
}

func uploadHandler(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
		return
	}

	// 1. Parse Form (Max 30MB)
	r.ParseMultipartForm(30 << 20)

	file, handler, err := r.FormFile("excelFile")
	if err != nil {
		http.Error(w, "No file uploaded", http.StatusBadRequest)
		return
	}
	defer file.Close()

	ext := strings.ToLower(filepath.Ext(handler.Filename))
	if ext != ".xlsx" && ext != ".xls" {
		http.Error(w, "Invalid file type", http.StatusBadRequest)
		return
	}

	var result [][]string
	var processErr error

	// 2. Strategi Percabangan Efisien
	if ext == ".xlsx" {
		// OPTIMISASI 1: XLSX tidak perlu simpan ke disk. Baca langsung dari stream memory.
		result, processErr = processXLSXStream(file)
	} else {
		// Untuk XLS (Lib C), kita terpaksa simpan ke disk dulu karena library butuh path file.
		filename := filepath.Join("uploads", fmt.Sprintf("temp-%s", handler.Filename))
		dst, err := os.Create(filename)
		if err != nil {
			http.Error(w, "Server Disk Error", http.StatusInternalServerError)
			return
		}
		// Copy data ke file
		if _, err := io.Copy(dst, file); err != nil {
			dst.Close()
			http.Error(w, "File Write Error", http.StatusInternalServerError)
			return
		}
		dst.Close()

		// Pastikan file dihapus setelah selesai
		defer os.Remove(filename)

		// Proses XLS
		result, processErr = processXLS(filename)
	}

	if processErr != nil {
		log.Printf("Processing error: %v", processErr)
		http.Error(w, "Error processing file", http.StatusInternalServerError)
		return
	}

	if len(result) == 0 {
		http.Error(w, `{"error": "No valid data found"}`, http.StatusBadRequest)
		return
	}

	// 3. Response JSON Cepat
	w.Header().Set("Content-Type", "application/json")
	// Menggunakan encoder goccy/go-json yang lebih cepat
	json.NewEncoder(w).Encode(result)
}

// =================================================================================
// LOGIC PROCESSOR (SINGLE PASS)
// =================================================================================

// SheetProcessor adalah state object untuk menyimpan posisi header per sheet
type SheetProcessor struct {
	headerFound bool
	headerMap   map[string]int // Key: "nopol" -> Val: 5 (index kolom)
	rowsBuffer  [][]string
}

func NewSheetProcessor() *SheetProcessor {
	return &SheetProcessor{
		headerFound: false,
		headerMap:   make(map[string]int),
	}
}

// ProcessRow dipanggil untuk SETIAP baris.
// Return true jika baris valid dan masuk result, false jika skip/header detection.
func (sp *SheetProcessor) ProcessRow(row []string) ([]string, bool) {
	// Skip baris kosong
	if len(row) == 0 {
		return nil, false
	}

	// Jika Header belum ketemu, coba cari di baris ini
	if !sp.headerFound {
		sp.detectHeader(row)
		return nil, false // Baris header tidak dimasukkan ke output data
	}

	// Jika Header SUDAH ketemu, langsung mapping menggunakan index (O(1) lookup)
	// Ini jauh lebih cepat daripada string matching berulang-ulang
	record := make(map[string]string, len(outputOrder))
	hasRequired := false

	for key, colIdx := range sp.headerMap {
		if colIdx < len(row) {
			val := row[colIdx]
			if val != "" {
				// Cleaning ringan
				val = strings.TrimSpace(val)

				// Optimisasi cleaning spesifik
				if key == "saldo" {
					val = strings.ReplaceAll(val, ",", ".")
					if num, err := strconv.ParseFloat(val, 64); err == nil {
						val = fmt.Sprintf("%.0f", math.Round(num))
					}
				} else if key == "nopol" {
					val = strings.ReplaceAll(val, " ", "")
					if val != "" {
						hasRequired = true
					}
				}
				record[key] = val
			}
		}
	}

	// Validasi Data
	if hasRequired {
		rowData := make([]string, len(outputOrder))
		for i, outKey := range outputOrder {
			rowData[i] = record[outKey] // Default string kosong "" jika map nil
		}
		return rowData, true
	}

	return nil, false
}

func (sp *SheetProcessor) detectHeader(row []string) {
	foundCount := 0
	tempMap := make(map[string]int)

	for colIdx, cell := range row {
		if cell == "" {
			continue
		}
		// Normalize cell header sekali saja
		cleanHeader := strings.ToLower(strings.ReplaceAll(cell, " ", ""))

		for _, mapping := range columnMappings {
			// Cek apakah sudah ketemu kolom ini sebelumnya di baris ini?
			if _, exists := tempMap[mapping.Key]; exists {
				continue
			}

			// Cek aliases
			for _, alias := range mapping.Aliases {
				if strings.Contains(cleanHeader, alias) {
					tempMap[mapping.Key] = colIdx
					foundCount++
					break // Lanjut ke mapping berikutnya
				}
			}
		}
	}

	// Kriteria Header Valid: Minimal ketemu kolom 'nopol' atau minimal 3 kolom cocok
	_, nopolExists := tempMap["nopol"]
	if nopolExists || foundCount >= 3 {
		sp.headerMap = tempMap
		sp.headerFound = true
	}
}

// =================================================================================
// READERS
// =================================================================================

// processXLSXStream membaca XLSX baris per baris (Streaming Iterator)
// Hemat Memory karena tidak load semua ke RAM
func processXLSXStream(reader io.Reader) ([][]string, error) {
	f, err := excelize.OpenReader(reader)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	var finalResult [][]string

	for _, sheetName := range f.GetSheetList() {
		// Gunakan Rows() iterator, bukan GetRows()
		rows, err := f.Rows(sheetName)
		if err != nil {
			continue
		}

		processor := NewSheetProcessor()
		rowCounter := 0

		for rows.Next() {
			rowCounter++
			// Batasi scan header hanya di 20 baris pertama agar CPU tidak habis scanning text di baris data
			if !processor.headerFound && rowCounter > 20 {
				break
			}

			// Ambil kolom mentah (string slice)
			cols, err := rows.Columns()
			if err != nil {
				continue
			}

			if data, ok := processor.ProcessRow(cols); ok {
				finalResult = append(finalResult, data)
			}
		}
		rows.Close()
	}

	return finalResult, nil
}

// processXLS membaca XLS baris per baris dan langsung memproses
func processXLS(filePath string) ([][]string, error) {
	wb, err := xls.OpenFile(filePath, "UTF-8")
	if err != nil {
		return nil, err
	}
	defer wb.Close()

	var finalResult [][]string

	// Batasi 20 sheet
	for i := 0; i < 20; i++ {
		sheet, err := wb.OpenWorkSheet(i)
		if err != nil {
			break
		}

		func() {
			defer sheet.Close()
			processor := NewSheetProcessor()

			// Loop Rows
			for rIdx, row := range sheet.Rows {
				// Stop scan header jika sampai baris 20 belum ketemu
				if !processor.headerFound && rIdx > 20 {
					break
				}

				// Konversi row.Cells ke []string
				strRow := make([]string, len(row.Cells))
				for cIdx, cell := range row.Cells {
					strRow[cIdx] = cell.Value.String()
				}

				if data, ok := processor.ProcessRow(strRow); ok {
					finalResult = append(finalResult, data)
				}
			}
		}()
	}

	return finalResult, nil
}
