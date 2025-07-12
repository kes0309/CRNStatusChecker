package main

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io/ioutil"
	"net/http"
	"os"
	"strings"
	"time"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/storage"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/widget"

	"github.com/xuri/excelize/v2"
)

type ResponseData struct {
	BNo     string `json:"b_no"`
	BStt    string `json:"b_stt"`
	TaxType string `json:"tax_type"`
	EndDt   string `json:"end_dt"`
}

func readServiceKey() string {
	data, err := os.ReadFile("serviceKey.txt")
	if err != nil {
		return ""
	}
	return strings.TrimSpace(string(data))
}

func saveServiceKey(key string) {
	_ = os.WriteFile("serviceKey.txt", []byte(key), 0644)
}

func checkCRNStatus(serviceKey, inputPath, outputPath string) string {
	f, err := excelize.OpenFile(inputPath)
	if err != nil {
		return "엑셀 파일을 열 수 없습니다"
	}
	rows, err := f.GetRows("Sheet1")
	if err != nil || len(rows) < 2 {
		return "엑셀 데이터가 유효하지 않습니다"
	}

	var crns []string
	for i, row := range rows {
		if i == 0 {
			if row[0] != "사업자등록번호" {
				return `"사업자등록번호" 열이 첫 행에 필요합니다`
			}
			continue
		}
		crns = append(crns, row[0])
	}

	url := "http://api.odcloud.kr/api/nts-businessman/v1/status?serviceKey=" + serviceKey
	var result [][]string

	for i := 0; i < len(crns); i += 100 {
		end := i + 100
		if end > len(crns) {
			end = len(crns)
		}
		chunk := crns[i:end]

		reqBody, _ := json.Marshal(map[string][]string{
			"b_no": chunk,
		})

		resp, err := http.Post(url, "application/json", bytes.NewBuffer(reqBody))
		if err != nil {
			return "인터넷 연결을 확인하세요"
		}
		defer resp.Body.Close()

		if resp.StatusCode != 200 {
			return fmt.Sprintf("오류: 상태코드 %d", resp.StatusCode)
		}

		body, _ := ioutil.ReadAll(resp.Body)
		var parsed map[string][]ResponseData
		json.Unmarshal(body, &parsed)

		for _, entry := range parsed["data"] {
			endDate := ""
			if entry.EndDt != "" {
				t, err := time.Parse("20060102", entry.EndDt)
				if err == nil {
					endDate = t.Format("2006-01-02")
				}
			}
			taxType := entry.TaxType
			if strings.Contains(taxType, "등록되지 않은") {
				taxType = "국세청에 등록되지 않음"
			}
			result = append(result, []string{entry.BNo, entry.BStt, taxType, endDate})
		}
	}

	outFile := excelize.NewFile()
	sheet := outFile.GetSheetName(0)
	headers := []string{"사업자등록번호", "납세자상태", "과세유형", "폐업일"}

	for i, h := range headers {
		cell, _ := excelize.CoordinatesToCellName(i+1, 1)
		outFile.SetCellValue(sheet, cell, h)
	}

	for i, row := range result {
		for j, val := range row {
			cell, _ := excelize.CoordinatesToCellName(j+1, i+2)
			outFile.SetCellValue(sheet, cell, val)
		}
	}

	err = outFile.SaveAs(outputPath + "/사업자상태.xlsx")
	if err != nil {
		return "엑셀 저장 오류: " + err.Error()
	}

	return ""
}

func main() {
	a := app.New()
	w := a.NewWindow("사업자 상태 확인")
	w.Resize(fyne.NewSize(800, 150))

	// 서비스 키
	serviceKeyEntry := widget.NewEntry()
	serviceKeyEntry.SetText(readServiceKey())

	saveKeyBtn := widget.NewButton("서비스키 저장", func() {
		key := strings.TrimSpace(serviceKeyEntry.Text)
		if key == "" {
			dialog.ShowInformation("오류", "서비스 키를 입력하세요", w)
			return
		}
		saveServiceKey(key)
		dialog.ShowInformation("알림", "저장되었습니다", w)
	})

	// 입력 엑셀
	inputLabel := widget.NewLabel("파일을 선택하세요")
	selectInputBtn := widget.NewButton("엑셀 불러오기", func() {
		fd := dialog.NewFileOpen(func(r fyne.URIReadCloser, err error) {
			if r != nil {
				inputLabel.SetText(r.URI().Path())
			}
		}, w)
		fd.SetFilter(storage.NewExtensionFileFilter([]string{".xlsx"}))
		fd.Show()
	})

	// 출력 폴더
	outputLabel := widget.NewLabel("저장 폴더 선택")
	selectOutputBtn := widget.NewButton("저장 폴더 선택", func() {
		dialog.NewFolderOpen(func(list fyne.ListableURI, err error) {
			if list != nil {
				outputLabel.SetText(list.Path())
			}
		}, w).Show()
	})

	// 조회 버튼
	checkBtn := widget.NewButton("사업자 상태 조회", func() {
		sk := strings.TrimSpace(serviceKeyEntry.Text)
		in := inputLabel.Text
		out := outputLabel.Text

		if sk == "" {
			dialog.ShowInformation("오류", "서비스 키를 입력하세요", w)
			return
		}
		if in == "" || out == "" {
			dialog.ShowInformation("오류", "입력 파일 및 저장 경로를 확인하세요", w)
			return
		}
		saveServiceKey(sk)
		errMsg := checkCRNStatus(sk, in, out)
		if errMsg == "" {
			dialog.ShowInformation("완료", "저장 위치: "+out+"/사업자상태.xlsx", w)
		} else {
			dialog.ShowError(fmt.Errorf(errMsg), w)
		}
	})

	// 레이아웃
	content := container.NewVBox(
		widget.NewLabel("서비스 키:"), serviceKeyEntry, saveKeyBtn,
		selectInputBtn, inputLabel,
		selectOutputBtn, outputLabel,
		checkBtn,
	)

	w.SetContent(content)
	w.ShowAndRun()
}
