package main

import (
	"context"
	"encoding/binary"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

type App struct {
	ctx context.Context
}

func NewApp() *App {
	return &App{}
}

func (a *App) startup(ctx context.Context) {
	a.ctx = ctx
}

func (a *App) ConvertX431ToXlsx(path string, endTimeStr string) (string, error) {
	if !strings.HasSuffix(strings.ToLower(path), ".x431") {
		return "", fmt.Errorf("invalid file type, expected .x431")
	}

	f, err := os.Open(path)
	if err != nil {
		return "", err
	}
	defer f.Close()

	buffer := make([]byte, 128)

	// Read channel count
	f.Seek(0x134, io.SeekStart)
	f.Read(buffer[:1])
	columnCount := int(buffer[0]) / 4

	// Skip to first length field
	f.Seek(0x0c, io.SeekStart)
	f.Read(buffer[:4])
	var32 := int32(binary.LittleEndian.Uint32(buffer[:4]))
	f.Seek(int64(var32), io.SeekCurrent)

	// Skip headers
	for i := 0; i < 8; i++ {
		f.Read(buffer[:2])
		var16 := int(binary.LittleEndian.Uint16(buffer[:2]))
		f.Seek(int64(var16-2), io.SeekCurrent)
	}

	// Read point values
	var pointValues []string
	fileSize := getFileSize(f)
	for {
		pos, _ := f.Seek(0, io.SeekCurrent)
		if pos >= fileSize {
			break
		}

		if _, err := f.Read(buffer[:2]); err != nil {
			break
		}
		var16 := int(binary.LittleEndian.Uint16(buffer[:2]))

		if var16 <= 2 {
			break
		}

		valBuffer := make([]byte, var16-2)
		if _, err := f.Read(valBuffer); err != nil {
			break
		}

		value := string(valBuffer[:len(valBuffer)-1]) // Skip null terminator
		pointValues = append(pointValues, value)
	}

	// Read column headers
	columnNames := make([]string, columnCount+1)
	columnNames[0] = "Time"

	f.Seek(0x138, io.SeekStart)
	for i := 0; i < columnCount; i++ {
		f.Read(buffer[:4])
		index := int(binary.LittleEndian.Uint16(buffer[:2]))
		if index != 0 && index-0x09 < len(pointValues) {
			columnNames[i+1] = fmt.Sprintf("%d. %s", i+1, pointValues[index-0x09])
		}
	}

	for i := 0; i < columnCount; i++ {
		f.Read(buffer[:4])
		index := int(binary.LittleEndian.Uint16(buffer[:2]))
		if index != 0 && index-0x09 < len(pointValues) {
			columnNames[i+1] = fmt.Sprintf("%s (%s)", columnNames[i+1], pointValues[index-0x09])
		}
	}

	// Prepare to read data
	f.Seek(0x11c, io.SeekStart)
	f.Read(buffer[:2])
	var16 := int(binary.LittleEndian.Uint16(buffer[:2]))

	f.Seek(int64(var16+8), io.SeekStart)
	f.Read(buffer[:8])
	recordsCount := int(binary.LittleEndian.Uint32(buffer[:4]))

	totalRows := (recordsCount / 4) / columnCount

	// Create Excel file
	sessionDate := parseSessionDate(path)
	ef := excelize.NewFile()
	defer ef.Close()

	var sheetName string
	if !sessionDate.IsZero() {
		sheetName = sessionDate.Format("02-01-06 1504")
	} else {
		sheetName = "Sheet1"
	}
	ef.SetSheetName("Sheet1", sheetName)

	// Column headers always start at row 1
	headerRow := 1

	// Write column headers
	for i, name := range columnNames {
		cell, _ := excelize.CoordinatesToCellName(i+1, headerRow)
		ef.SetCellValue(sheetName, cell, name)
	}

	// Calculate custom sampling interval if endTimeStr is provided
	samplingInterval := 1.0
	useCustomSampling := false
	if endTimeStr != "" && !sessionDate.IsZero() {
		endTime, err := time.Parse("2006-01-02T15:04", endTimeStr)
		if err != nil {
			// Fallback for some browsers that might include seconds
			endTime, err = time.Parse("2006-01-02T15:04:05", endTimeStr)
		}

		if err == nil {
			if !endTime.After(sessionDate) {
				return "", fmt.Errorf("end time must be after start time")
			}
			duration := endTime.Sub(sessionDate)
			samplingInterval = duration.Seconds() / float64(totalRows)
			useCustomSampling = true
		}
	}

	// Pre-create a datetime style for the Time column (when session date is known)
	timeStyleID := 0
	if !sessionDate.IsZero() {
		style, err := ef.NewStyle(&excelize.Style{
			NumFmt: 22, // built-in "m/d/yy h:mm" — Excel will show full datetime
		})
		if err == nil {
			timeStyleID = style
		}
	}

	// Write data rows
	for i := 0; i < totalRows; i++ {
		rowIdx := headerRow + 1 + i
		timeCell, _ := excelize.CoordinatesToCellName(1, rowIdx)
		if !sessionDate.IsZero() {
			var ts time.Time
			if useCustomSampling {
				ts = sessionDate.Add(time.Duration(float64(i)*samplingInterval*float64(time.Second)))
			} else {
				ts = sessionDate.Add(time.Duration(i) * time.Second)
			}
			ef.SetCellValue(sheetName, timeCell, ts)
			if timeStyleID != 0 {
				ef.SetCellStyle(sheetName, timeCell, timeCell, timeStyleID)
			}
		} else {
			ef.SetCellValue(sheetName, timeCell, i+1)
		}

		for j := 0; j < columnCount; j++ {
			f.Read(buffer[:4])
			index := int(binary.LittleEndian.Uint16(buffer[:2])) - 0x09
			cell, _ := excelize.CoordinatesToCellName(j+2, rowIdx)
			if index >= 0 && index < len(pointValues) {
				setNumericOrString(ef, sheetName, cell, pointValues[index])
			} else {
				ef.SetCellValue(sheetName, cell, 0)
			}
		}
	}

	outputFile := strings.TrimSuffix(path, filepath.Ext(path)) + ".xlsx"
	if err := ef.SaveAs(outputFile); err != nil {
		return "", err
	}

	return outputFile, nil
}

func getFileSize(f *os.File) int64 {
	stat, _ := f.Stat()
	return stat.Size()
}

// parseSessionDate extracts a YYYYMMDDHHmmss datetime from the filename.
// For example "VENDOR_988784046340_20250709161906-notes.x431" → 2025-07-09 16:19:06.
// Returns zero time if no valid datetime is found.
func parseSessionDate(path string) time.Time {
	base := strings.TrimSuffix(filepath.Base(path), filepath.Ext(path))
	// Match a 14-digit block: YYYYMMDDHHmmss
	re := regexp.MustCompile(`(\d{14})`)
	if m := re.FindString(base); m != "" {
		t, err := time.ParseInLocation("20060102150405", m, time.Local)
		if err == nil {
			return t
		}
	}
	return time.Time{}
}

// setNumericOrString writes the value as float64 if it looks like a number,
// otherwise as a plain string. This makes Excel treat the cell as a number
// regardless of the system's decimal-separator locale setting.
func setNumericOrString(ef *excelize.File, sheet, cell, value string) {
	if f, err := strconv.ParseFloat(strings.TrimSpace(value), 64); err == nil {
		ef.SetCellValue(sheet, cell, f)
	} else {
		ef.SetCellValue(sheet, cell, value)
	}
}
