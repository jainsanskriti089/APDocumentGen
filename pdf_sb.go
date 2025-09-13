package main

import (
	"bufio"
	"encoding/csv"
	"flag"
	"fmt"
	"log"
	"os"
	"sort"
	"strconv"
	"strings"
	"time"

	// "path/filepath"
	"github.com/jung-kurt/gofpdf"
	"github.com/shakinm/xlsReader/xls"
	"github.com/xuri/excelize/v2"
)

type SBData struct {
	apptTime    string //A5
	ptName      string //E5
	ptID        string //D5
	ptType      string
	dob         string //H5
	apptType    string
	apptDate    string //C1
	apptStatus  string //L5
	wccAge      string
	visitReason string
	isWcc       bool
	isNP        bool
	isNV        bool
	ins         string //Column N
	inNtwk      string //Column F
	copay       string //Column G or H (sick)
	coins       string //Column I
	ptBalance   string //column J
	sibBalance  string // Column K
	insCard     string //Column E
}

type VaccineDetails struct {
	site         string
	route        string
	brandName    string
	manufacturer string
	lotNumber    string
	expiryDt     string
	visDt        string
	qty          string
}
type BFdata struct {
	numOfPages        int
	numOfHandoutPages int
}

var bfData = make(map[string]BFdata)
var immData = make(map[string]string)
var vacData = make(map[string]VaccineDetails)

func main() {

	filterPtr := flag.String("type", "", "Appointment Filter Type")
	fromTimePtr := flag.String("from", "", "Appointment From")
	toTimePtr := flag.String("to", "", "Appointment To")
	flag.Parse()

	immData["2-5D"] = ""
	immData["1M"] = ""
	immData["2M"] = "HepB,Rota,DT/DTaP,Hib,PCV,IPV"
	immData["4M"] = "Rota,DT/DTaP,Hib,PCV,IPV"
	immData["6M"] = "HepB,Rota,DT/DTaP,Hib,PCV,IPV"
	immData["9M"] = ""
	immData["12M"] = "MMR,VZV,HepA"
	immData["15M"] = "DT/DTaP,IPV,Hib,PCV"
	immData["18M"] = "HepA"
	immData["2Y"] = ""
	immData["2.5Y"] = ""
	immData["3Y"] = ""
	immData["4Y"] = "DTaP,Polio,MMR,VZV"
	immData["5Y"] = ""
	immData["6Y"] = ""
	immData["7Y"] = ""
	immData["8Y"] = ""
	immData["9Y"] = ""
	immData["10Y"] = ""
	immData["11Y"] = "Tdap,HPV,MCV"
	immData["12Y"] = ""
	immData["13Y"] = ""
	immData["14Y"] = ""
	immData["15Y"] = ""
	immData["16Y"] = "MenB,MCV"
	immData["17Y"] = ""
	immData["18Y"] = ""
	immData["19Y"] = ""
	immData["20Y"] = ""
	immData["21Y"] = ""

	bfData["2-5D"] = BFdata{numOfPages: 1, numOfHandoutPages: 1} // correct
	bfData["1M"] = BFdata{numOfPages: 1, numOfHandoutPages: 1}   // correct
	bfData["2M"] = BFdata{numOfPages: 2, numOfHandoutPages: 1}   // correct
	bfData["4M"] = BFdata{numOfPages: 2, numOfHandoutPages: 1}   // correct
	bfData["6M"] = BFdata{numOfPages: 2, numOfHandoutPages: 1}   // correct
	bfData["9M"] = BFdata{numOfPages: 2, numOfHandoutPages: 1}   // correct
	bfData["12M"] = BFdata{numOfPages: 2, numOfHandoutPages: 1}  // correct
	bfData["15M"] = BFdata{numOfPages: 2, numOfHandoutPages: 1}  // correct
	bfData["18M"] = BFdata{numOfPages: 3, numOfHandoutPages: 1}  // correct
	bfData["2Y"] = BFdata{numOfPages: 3, numOfHandoutPages: 1}   // correct
	bfData["2.5Y"] = BFdata{numOfPages: 2, numOfHandoutPages: 1} // correct
	bfData["3Y"] = BFdata{numOfPages: 2, numOfHandoutPages: 1}   // correct
	bfData["4Y"] = BFdata{numOfPages: 2, numOfHandoutPages: 1}   // correct
	bfData["5Y"] = BFdata{numOfPages: 2, numOfHandoutPages: 1}   // correct
	bfData["6Y"] = BFdata{numOfPages: 1, numOfHandoutPages: 1}   // correct
	bfData["7Y"] = BFdata{numOfPages: 1, numOfHandoutPages: 1}   // correct
	bfData["8Y"] = BFdata{numOfPages: 1, numOfHandoutPages: 1}   // correct
	bfData["9Y"] = BFdata{numOfPages: 1, numOfHandoutPages: 1}   // correct
	bfData["10Y"] = BFdata{numOfPages: 1, numOfHandoutPages: 1}  // correct
	bfData["11Y"] = BFdata{numOfPages: 3, numOfHandoutPages: 1}  // correct
	bfData["12Y"] = BFdata{numOfPages: 5, numOfHandoutPages: 1}  // correct
	bfData["13Y"] = BFdata{numOfPages: 5, numOfHandoutPages: 1}  // correct
	bfData["14Y"] = BFdata{numOfPages: 5, numOfHandoutPages: 1}  // correct
	bfData["15Y"] = BFdata{numOfPages: 3, numOfHandoutPages: 1}  // correct
	bfData["16Y"] = BFdata{numOfPages: 3, numOfHandoutPages: 1}  // correct
	bfData["17Y"] = BFdata{numOfPages: 3, numOfHandoutPages: 1}  // correct
	bfData["18Y"] = BFdata{numOfPages: 3, numOfHandoutPages: 1}  // correct
	bfData["19Y"] = BFdata{numOfPages: 3}
	bfData["20Y"] = BFdata{numOfPages: 3}
	bfData["21Y"] = BFdata{numOfPages: 3}

	_, _ = createDOSDir(pathAppt())
	loadVaccineData()
	f, patientList := readCreatedSBsFile(pathAppt())
	appointments, apptList := loadApptCSV(pathAppt())
	apptList = loadEligibilityReport("./Eligibility_Report.xlsx", "sheet1", appointments, apptList)

	makeDailyFolders(appointments, apptList)

	if *filterPtr == "Same" || *filterPtr == "E-Consult" {
		fmt.Printf("Inside filtered section ")
		pdf := gofpdf.New("P", "mm", "Letter", "")
		pdf = createFilteredPdf(pdf, appointments, *filterPtr, *fromTimePtr, *toTimePtr, apptList)
		if pdf.Err() {
			log.Fatalf("Failed creating PDF report: %s\n", pdf.Error())
		}
		err := savePDF(pdf, apptList[0].apptDate, *filterPtr)
		if err != nil {
			log.Fatalf("Cannot save PDF: %s|n", err)
		}
		return

	}
	// creates new fpdf instance with portrait mode P with millimeter unit, Letter specifies page size

	change2DOSFolder(appointments, apptList)
	createConfirmedPdf(appointments, apptList)
	f.Close()
	_ = os.Chdir("..")
	f, patientList = readCreatedSBsFile(pathAppt())

	createUnconfirmed := false
	change2DOSFolder(appointments, apptList)
	if createUnconfirmed {
		createUnconfirmedPdf(appointments, apptList, f, patientList)
	}

	_ = os.Chdir("..")
	createCovidPdf(appointments, apptList, f, patientList)
	_ = os.Chdir("..")
	createEmptyIZAdminPdf(appointments, apptList, f, patientList)
	_ = os.Chdir("..")
	createVacRefusalPdf(appointments, apptList, f, patientList)
	_ = os.Chdir("..")
	f.Close()
}

func change2DOSFolder(appointments int, apptList [200]SBData) {
	apptDate := ""

	for i := 0; i < appointments; i++ {
		apptData := apptList[i]
		if apptData.apptDate != "" && apptDate == "" {
			apptDate = strings.ReplaceAll(apptData.apptDate, "/", "-")
			err := os.Mkdir(apptDate, 0777)
			if err != nil && !os.IsExist(err) {
				log.Fatal(err)
			}

			err = os.Chdir(apptDate)
			if err != nil {
				log.Fatal(err)
			}
			break
		}
	}

	return
}

func makeDailyFolders(appointments int, apptList [200]SBData) {

	change2DOSFolder(appointments, apptList)
	// curDir,_ := os.Getwd()
	// fmt.Printf("Inside directory makeDailyFolders [%s] \n ", curDir)
	for i := 0; i < appointments; i++ {
		apptData := apptList[i]
		if apptData.ptName != "" && !strings.Contains(apptData.ptName, "Blocked, Appt") && !strings.Contains(apptData.visitReason, "Nurse Visit - Non Billable") {
			err := os.Mkdir(
				strings.ReplaceAll(strings.TrimSpace(apptData.apptTime), ":", "-")+"-"+
					strings.ReplaceAll(strings.ReplaceAll(strings.TrimSpace(apptData.ptName), ",", "-"), " ", "-"),
				0777)
			if err != nil && !os.IsExist(err) {
				log.Fatal(err)
			}
		}

	}
	os.Chdir("..")
}

func createConfirmedPdf(appointments int, apptList [200]SBData) {
	var izAdminPageCnt int

	data := loadCSV(path())
	for i := 0; i < appointments; i++ {
		ptPdf := gofpdf.New("P", "mm", "Letter", "")
		apptData := apptList[i]
		if strings.Contains(apptData.ptName, "Blocked, Appt") || strings.Contains(apptData.visitReason, "Nurse Visit - Non Billable") {
			continue
		}
		ptPdf = newReport(ptPdf, apptData)
		ptPdf = header(ptPdf, data[0])
		ptPdf = table(ptPdf, data[1:])
		ptPdf = sbTableFront(ptPdf, apptData)
		// till this point total # of pages for the patient is 2
		totalPtPages := 2

		if apptData.isNP || (apptData.isWcc && strings.ToUpper(apptData.wccAge) == "2-5D") {
			ptPdf = initialHxImages(ptPdf, apptData)
			totalPtPages += 2
		}
		if apptData.isWcc {
			if apptData.wccAge != "" {
				// add audiogram
				izAdminPageCnt, ptPdf = addAudio(ptPdf, apptData)
				totalPtPages += izAdminPageCnt

				// add izadmin
				izAdminPageCnt, ptPdf = addIzAdmin(ptPdf, apptData)
				totalPtPages += izAdminPageCnt

				/* add bright futures pages */
				ptPdf = otherImages(ptPdf, apptData, bfData[apptData.wccAge].numOfPages, "P", true)
				totalPtPages += bfData[apptData.wccAge].numOfPages
				if (totalPtPages % 2) != 0 {
					ptPdf.AddPage()
					totalPtPages += 1
				}
				/* add bright futures handout pages */
				ptPdf = otherImages(ptPdf, apptData, bfData[apptData.wccAge].numOfHandoutPages, "H", false)
				totalPtPages += bfData[apptData.wccAge].numOfHandoutPages
				if (totalPtPages % 2) != 0 {
					ptPdf.AddPage()
					totalPtPages += 1
				}
			}
		} else if apptData.isNV {
			if apptData.wccAge != "" {
				izAdminPageCnt, ptPdf = addIzAdmin(ptPdf, apptData)
			} else {
				rsn := strings.ToLower(apptData.visitReason)
				// fmt.Print(rsn)
				if strings.Contains(rsn, "tb") && !strings.Contains(rsn, "reading") {
					// fmt.Print("tb test")
					izAdminPageCnt, ptPdf = addTB(ptPdf, apptData)
					totalPtPages += izAdminPageCnt
				}
				if strings.Contains(rsn, "hearing") || strings.Contains(rsn, "audio") {
					// fmt.Print("tb test")
					izAdminPageCnt, ptPdf = addAudio(ptPdf, apptData)
					totalPtPages += izAdminPageCnt
				}
				izAdminPageCnt, ptPdf = addNurseIzAdmin(ptPdf, apptData)
			}
			totalPtPages += izAdminPageCnt
		} else {
			izAdminPageCnt, ptPdf = addIzAdmin(ptPdf, apptData)
			totalPtPages += izAdminPageCnt
		}

		if (totalPtPages % 2) != 0 {
			ptPdf.AddPage()
		}
		if ptPdf.Err() {
			log.Fatalf("Failed creating Patient PDF report: %s\n", ptPdf.Error())
		}
		err := savePtPDF(ptPdf, apptData.ptName, apptData.apptTime, apptList[0].apptDate, "confirmed", "SB-")
		if err != nil {
			log.Fatalf("Cannot save Patient PDF: %s|n", err)
		}
	}
	return
}

func createUnconfirmedPdf(appointments int, apptList [200]SBData, writeFile *os.File, patientList [200]string) {
	var izAdminPageCnt int
	var patientMap map[string]bool

	fmt.Printf("Inside createUnconfirmedPDF \n")
	// initialize map
	patientMap = make(map[string]bool)
	for _, s := range patientList {
		patientMap[s] = true
	}

	data := loadCSV(path())
	for i := 0; i < appointments; i++ {
		ptPdf := gofpdf.New("P", "mm", "Letter", "")
		apptData := apptList[i]
		// fmt.Printf("Looking at patient %s with appt status %s \n",apptData.ptName,  apptData.apptStatus)
		if strings.Contains(apptData.ptName, "Blocked, Appt") || strings.Contains(apptData.apptStatus, "Confirmed") || strings.Contains(apptData.visitReason, "Nurse Visit - Non Billable") {
			continue
		}
		if _, presence := patientMap[apptData.ptName]; presence {
			fmt.Printf("found patient %s, skipping ...\n", apptData.ptName)
			continue
		}
		writeFile.Seek(0, os.SEEK_END)
		_, err := writeFile.WriteString(apptData.ptName)
		if err != nil {
			log.Fatalf("Failed to write  '%s': %s\n", apptData.ptName, err.Error())
		}
		// fmt.Printf("Wrote %d bytes for %s \n", n, apptData.ptName)
		_, err = writeFile.WriteString("\n")
		if err != nil {
			log.Fatalf("Failed to write  '%s': %s\n", apptData.ptName, err.Error())
		}
		ptPdf = newReport(ptPdf, apptData)
		ptPdf = header(ptPdf, data[0])
		ptPdf = table(ptPdf, data[1:])
		// ptPdf = image(ptPdf)
		ptPdf = sbTableFront(ptPdf, apptData)
		// till this point total # of pages for the patient is 2
		totalPtPages := 2

		if apptData.isNP || (apptData.isWcc && strings.ToUpper(apptData.wccAge) == "2-5D") {
			if apptData.ptType == "NP" || (apptData.isWcc && strings.ToUpper(apptData.wccAge) == "2-5D") {
				ptPdf = initialHxImages(ptPdf, apptData)
				totalPtPages += 2
			}
		}

		if apptData.isWcc {
			if apptData.wccAge != "" {
				if !apptData.isNP && apptData.apptStatus == "Confirmed" {
					/* add bright futures pages */
					ptPdf = otherImages(ptPdf, apptData, bfData[apptData.wccAge].numOfPages, "P", true)
					totalPtPages += bfData[apptData.wccAge].numOfPages
					if (totalPtPages % 2) != 0 {
						ptPdf.AddPage()
						totalPtPages += 1
					}
					/* add bright futures handout pages */
					ptPdf = otherImages(ptPdf, apptData, bfData[apptData.wccAge].numOfHandoutPages, "H", false)
					totalPtPages += bfData[apptData.wccAge].numOfHandoutPages
					if (totalPtPages % 2) != 0 {
						ptPdf.AddPage()
						totalPtPages += 1
					}
					// add audiogram
					izAdminPageCnt, ptPdf = addAudio(ptPdf, apptData)
					totalPtPages += izAdminPageCnt

					// add izadmin
					izAdminPageCnt, ptPdf = addIzAdmin(ptPdf, apptData)
					totalPtPages += izAdminPageCnt
				} else {
					/* add bright futures pages */
					ptPdf = otherImages(ptPdf, apptData, bfData[apptData.wccAge].numOfPages, "P", true)
					totalPtPages += bfData[apptData.wccAge].numOfPages
					if (totalPtPages % 2) != 0 {
						ptPdf.AddPage()
						totalPtPages += 1
					}
					/* add bright futures handout pages */
					ptPdf = otherImages(ptPdf, apptData, bfData[apptData.wccAge].numOfHandoutPages, "H", false)
					totalPtPages += bfData[apptData.wccAge].numOfHandoutPages
					if (totalPtPages % 2) != 0 {
						ptPdf.AddPage()
						totalPtPages += 1
					}
					// add audiogram
					izAdminPageCnt, ptPdf = addAudio(ptPdf, apptData)
					totalPtPages += izAdminPageCnt

					// add izadmin
					izAdminPageCnt, ptPdf = addIzAdmin(ptPdf, apptData)
					totalPtPages += izAdminPageCnt
				}
			}
		} else if apptData.isNV {
			if apptData.wccAge != "" {
				izAdminPageCnt, ptPdf = addIzAdmin(ptPdf, apptData)
				totalPtPages += izAdminPageCnt
			} else {
				rsn := strings.ToLower(apptData.visitReason)
				if strings.Contains(rsn, "tb") && !strings.Contains(rsn, "reading") {
					fmt.Print("tb reading")
					izAdminPageCnt, ptPdf = addTB(ptPdf, apptData)
					totalPtPages += izAdminPageCnt
				}
				izAdminPageCnt, ptPdf = addNurseIzAdmin(ptPdf, apptData)
				totalPtPages += izAdminPageCnt
			}
		} else {
			izAdminPageCnt, ptPdf = addIzAdmin(ptPdf, apptData)
			totalPtPages += izAdminPageCnt
		}

		if (totalPtPages % 2) != 0 {
			ptPdf.AddPage()
		}

		if ptPdf.Err() {
			log.Fatalf("Failed creating Patient PDF report: %s\n", ptPdf.Error())
		}
		err = savePtPDF(ptPdf, apptData.ptName, apptData.apptTime, apptList[0].apptDate, "unconfirmed", "SB-")
		// fmt.Printf("Looking at patient %s with appt status %s \n",apptData.ptName,  apptData.apptStatus)
		if err != nil {
			log.Fatalf("Cannot save Patient PDF: %s|n", err)
		}
	}
	return
}

func createCovidPdf(appointments int, apptList [200]SBData, writeFile *os.File, patientList [200]string) {
	var patientMap map[string]bool

	// initialize map
	patientMap = make(map[string]bool)
	for _, s := range patientList {
		patientMap[s] = true
	}

	change2DOSFolder(appointments, apptList)

	for i := 0; i < appointments; i++ {
		ptPdf := gofpdf.New("P", "mm", "Letter", "")
		apptData := apptList[i]
		// fmt.Printf("Looking at patient %s with appt status %s \n",apptData.ptName,  apptData.apptStatus)
		if strings.Contains(apptData.ptName, "Blocked, Appt") || strings.Contains(apptData.visitReason, "Nurse Visit - Non Billable") {
			continue
		}
		apptStr := fmt.Sprintf("Name:%1.30s DOB: %s DOS: %s", apptData.ptName, apptData.dob, apptData.apptDate)
		ptPdf.AddPage()
		ptPdf.SetFont("Times", "", 12)
		ptPdf.SetFontSize(14)
		ptPdf.CellFormat(40, 7, apptStr, "0", 0, "L", false, 0, "")
		ptPdf.Ln(1)
		fileName := fmt.Sprintf("../BF/CovidP1.png")
		ptPdf.ImageOptions(fileName, 10, 20, -1, -1, false, gofpdf.ImageOptions{ImageType: "PNG", ReadDpi: true}, 0, "")
		ptPdf.AddPage()
		ptPdf.AddPage()
		ptPdf.CellFormat(40, 7, apptStr, "0", 0, "L", false, 0, "")
		ptPdf.Ln(1)
		fileName = fmt.Sprintf("../BF/CovidP2.png")
		ptPdf.ImageOptions(fileName, 10, 20, -1, -1, false, gofpdf.ImageOptions{ImageType: "PNG", ReadDpi: true}, 0, "")

		ptPdf.AddPage()
		/*
			ptPdf.AddPage()
			ptPdf.CellFormat(15, 7, apptStr, "0", 0, "", false, 0, "")
			ptPdf.Ln(1)
			fileName = fmt.Sprintf("../BF/IZAdminHeader.png")
			ptPdf.ImageOptions(fileName, 10, 20, -1, -1, false, gofpdf.ImageOptions{ImageType: "PNG", ReadDpi: true}, 0, "")
		*/
		ptPdf.AddPage()
		ptPdf.SetFont("Times", "", 12)
		ptPdf.SetFontSize(14)
		ptPdf.SetFillColor(235, 235, 235)
		ptPdf.CellFormat(192, 8, "VACCINE ADMINISTRATION RECORD", "", 1, "CB", true, 0, "")
		ptPdf.CellFormat(192, 8, "Arti Pediatrics, Inc", "", 1, "CB", true, 0, "")
		ptPdf.CellFormat(192, 8, "2500 Hospital Dr, Suite 8A", "", 1, "CB", true, 0, "")
		ptPdf.CellFormat(192, 8, "Mountain VIew, CA 94040", "", 1, "CB", true, 0, "")
		ptPdf.CellFormat(192, 8, "408-462-9261(0) / 408-701-5006 (Fax)", "", 1, "CB", true, 0, "")

		// pdf.SetFontSize(12)
		// apptStr := fmt.Sprintf("Name:%1.30s DOB: %s DOS: %s", apptData.ptName, apptData.dob, apptData.apptDate)
		// pdf.CellFormat(15, 7, apptStr, "0", 0, "", false, 0, "")
		// pdf.Ln(1)
		ptPdf.CellFormat(192, 7, "", "", 1, "", false, 0, "")
		ptPdf.Ln(1)
		ptPdf.SetFontSize(12)
		apptStr = fmt.Sprintf("Patient Name     : %s ", apptData.ptName)
		ptPdf.CellFormat(120, 7, apptStr, "B", 1, "", false, 0, "")
		apptStr = fmt.Sprintf("Birth Date          : %s ", apptData.dob)
		ptPdf.CellFormat(120, 7, apptStr, "B", 1, "", false, 0, "")
		apptStr = fmt.Sprintf("Date of Service  : %s ", apptData.apptDate)
		ptPdf.CellFormat(120, 7, apptStr, "B", 1, "", false, 0, "")
		ptPdf.SetFontSize(9)
		ptPdf.CellFormat(192, 7, "A copy of the appropriate Centers for Disease Control and Prevention Vaccine Information Statement was provided to me. By signing below, I agree that", "", 1, "", false, 0, "")
		// pdf.CellFormat(192, 7, "By signing below, I agree that", "", 1, "", false, 0, "")
		ptPdf.CellFormat(192, 7, "* I have read or had explained to me the information about this disease and the vaccine.", "", 1, "", false, 0, "")
		ptPdf.CellFormat(192, 7, "* I had an opportunity to ask questions, and those questions were answered satisfactorily.", "", 1, "T", false, 0, "")
		ptPdf.CellFormat(192, 7, "* I believe that I understand the benefits and risks of the vaccine.", "", 1, "T", false, 0, "")
		ptPdf.CellFormat(192, 7, "* I ask that the vaccine be given to me or the person named above (for whom I am authorized to make the request).", "", 1, "T", false, 0, "")
		ptPdf.CellFormat(192, 7, "Every time I initial the \"Parent/Guardian/Patient Initials box, I agree that all of these actions have occured for the vaccine listed in that row.", "", 1, "", false, 0, "")
		ptPdf.CellFormat(192, 12, "", "", 1, "", false, 0, "")
		ptPdf.Ln(1)
		ptPdf.CellFormat(120, 7, "Parent/Guardian/Patient Signature                                                Date", "T", 1, "", false, 0, "")
		// fileName := fmt.Sprintf("../BF/IZAdminHeader.png")
		// pdf.ImageOptions(fileName, 10, 20, -1, -1, false, gofpdf.ImageOptions{ImageType: "PNG", ReadDpi: true}, 0, "")

		// ptPdf.SetXY(90, 90)
		ptPdf.Ln(1)
		// fmt.Printf("length of vaccineForWcc %d\n", len(vaccineForWcc))
		ptPdf.SetFontSize(9)
		ptPdf = izTableHeader(ptPdf)
		ptPdf.SetFillColor(255, 255, 255)
		vaccineForWcc := [6]string{"Covid", "Flu", " ", " ", " ", " "}
		for _, vaccine := range vaccineForWcc {
			// fmt.Printf("ptName [%s] wccAge [%s] vaccine [%s]\n", apptData.ptName, wccAge, vaccine)
			for i := 0; i < 12; i++ {
				switch i {
				case 0:
					ptPdf.CellFormat(18, 7, vaccine, "1", 0, "", true, 0, "")
				case 1:
					if vaccine != " " {
						ptPdf.CellFormat(18, 7, apptData.apptDate, "1", 0, "", true, 0, "")
					} else {
						ptPdf.CellFormat(18, 7, " ", "1", 0, "", true, 0, "")
					}
				case 2:
					ptPdf.CellFormat(12, 7, " ", "1", 0, "C", true, 0, "")
				case 3:
					ptPdf.CellFormat(12, 7, vacData[strings.ReplaceAll(vaccine, " ", "")].route, "1", 0, "C", true, 0, "")
				case 4:
					ptPdf.CellFormat(18, 7, vacData[strings.ReplaceAll(vaccine, " ", "")].brandName, "1", 0, "C", true, 0, "")
				case 5:
					ptPdf.CellFormat(18, 7, vacData[strings.ReplaceAll(vaccine, " ", "")].manufacturer, "1", 0, "C", true, 0, "")
				case 6:
					ptPdf.CellFormat(18, 7, vacData[strings.ReplaceAll(vaccine, " ", "")].lotNumber, "1", 0, "C", true, 0, "")
				case 7:
					ptPdf.CellFormat(18, 7, vacData[strings.ReplaceAll(vaccine, " ", "")].expiryDt, "1", 0, "C", true, 0, "")
				case 8:
					ptPdf.CellFormat(16, 7, vacData[strings.ReplaceAll(vaccine, " ", "")].visDt, "1", 0, "C", true, 0, "")
				case 9:
					if vaccine != " " {
						ptPdf.CellFormat(17, 7, apptData.apptDate, "1", 0, "", true, 0, "")
					} else {
						ptPdf.CellFormat(17, 7, " ", "1", 0, "", true, 0, "")
					}
					// ptPdf.CellFormat(17, 7, apptData.apptDate,"1", 0, "", true, 0, "")
				case 10:
					ptPdf.CellFormat(12, 7, "", "1", 0, "", true, 0, "")
				case 11:
					ptPdf.CellFormat(18, 7, "", "1", 1, "", true, 0, "")
				}
			}
		}

		err := savePtPDF(ptPdf, apptData.ptName, apptData.apptTime, apptList[0].apptDate, "covid", "PW-")
		// fmt.Printf("Looking at patient %s with appt status %s \n",apptData.ptName,  apptData.apptStatus)
		if err != nil {
			log.Fatalf("Covid 3 - Cannot save Patient PDF: %s|n", err)
		}
	}
	return
}

func createEmptyIZAdminPdf(appointments int, apptList [200]SBData, writeFile *os.File, patientList [200]string) {
	var patientMap map[string]bool

	// initialize map
	patientMap = make(map[string]bool)
	for _, s := range patientList {
		patientMap[s] = true
	}

	change2DOSFolder(appointments, apptList)

	for i := 0; i < appointments; i++ {
		ptPdf := gofpdf.New("P", "mm", "Letter", "")
		apptData := apptList[i]
		// fmt.Printf("Looking at patient %s with appt status %s \n",apptData.ptName,  apptData.apptStatus)
		if strings.Contains(apptData.ptName, "Blocked, Appt") || strings.Contains(apptData.visitReason, "Nurse Visit - Non Billable") {
			continue
		}
		apptStr := fmt.Sprintf("Name:%1.30s DOB: %s DOS: %s", apptData.ptName, apptData.dob, apptData.apptDate)
		/*
			ptPdf.AddPage()
			ptPdf.SetFont("Times", "", 12)
			ptPdf.SetFontSize(14)
			ptPdf.CellFormat(15, 7, apptStr, "0", 0, "", false, 0, "")
			ptPdf.Ln(1)
			fileName := fmt.Sprintf("../BF/IZAdminHeader.png")
			ptPdf.ImageOptions(fileName, 10, 20, -1, -1, false, gofpdf.ImageOptions{ImageType: "PNG", ReadDpi: true}, 0, "")
		*/

		ptPdf.AddPage()
		ptPdf.SetFont("Times", "", 12)
		ptPdf.SetFontSize(14)
		ptPdf.SetFillColor(235, 235, 235)
		ptPdf.CellFormat(192, 8, "VACCINE ADMINISTRATION RECORD", "", 1, "CT", true, 0, "")
		ptPdf.CellFormat(192, 8, "Arti Pediatrics, Inc", "", 1, "CT", true, 0, "")
		ptPdf.CellFormat(192, 8, "2500 Hospital Dr, Suite 8A", "", 1, "CT", true, 0, "")
		ptPdf.CellFormat(192, 8, "Mountain VIew, CA 94040", "", 1, "CT", true, 0, "")
		ptPdf.CellFormat(192, 8, "408-462-9261(0) / 408-701-5006 (Fax)", "", 1, "CT", true, 0, "")

		// pdf.SetFontSize(12)
		// apptStr := fmt.Sprintf("Name:%1.30s DOB: %s DOS: %s", apptData.ptName, apptData.dob, apptData.apptDate)
		// pdf.CellFormat(15, 7, apptStr, "0", 0, "", false, 0, "")
		// pdf.Ln(1)
		ptPdf.CellFormat(192, 7, "", "", 1, "", false, 0, "")
		ptPdf.Ln(1)
		ptPdf.SetFontSize(12)
		apptStr = fmt.Sprintf("Patient Name   : %s ", apptData.ptName)
		ptPdf.CellFormat(120, 7, apptStr, "B", 1, "", false, 0, "")
		apptStr = fmt.Sprintf("Birth Date       : %s ", apptData.dob)
		ptPdf.CellFormat(120, 7, apptStr, "B", 1, "", false, 0, "")
		apptStr = fmt.Sprintf("Date of Service : %s ", apptData.apptDate)
		ptPdf.CellFormat(120, 7, apptStr, "B", 1, "", false, 0, "")
		ptPdf.SetFontSize(9)
		ptPdf.CellFormat(192, 7, "A copy of the appropriate Centers for Disease Control and Prevention Vaccine Information Statement was provided to me. By signing below, I agree that", "", 1, "", false, 0, "")
		// pdf.CellFormat(192, 7, "By signing below, I agree that", "", 1, "", false, 0, "")
		ptPdf.CellFormat(192, 7, "* I have read or had explained to me the information about this disease and the vaccine.", "", 1, "", false, 0, "")
		ptPdf.CellFormat(192, 7, "* I had an opportunity to ask questions, and those questions were answered satisfactorily.", "", 1, "T", false, 0, "")
		ptPdf.CellFormat(192, 7, "* I believe that I understand the benefits and risks of the vaccine.", "", 1, "T", false, 0, "")
		ptPdf.CellFormat(192, 7, "* I ask that the vaccine be given to me or the person named above (for whom I am authorized to make the request).", "", 1, "T", false, 0, "")
		ptPdf.CellFormat(192, 7, "Every time I initial the \"Parent/Guardian/Patient Initials box, I agree that all of these actions have occured for the vaccine listed in that row.", "", 1, "", false, 0, "")
		ptPdf.CellFormat(192, 12, "", "", 1, "", false, 0, "")
		ptPdf.Ln(1)
		ptPdf.CellFormat(120, 7, "Parent/Guardian/Patient Signature                                                Date", "T", 1, "", false, 0, "")
		// fileName := fmt.Sprintf("../BF/IZAdminHeader.png")
		// pdf.ImageOptions(fileName, 10, 20, -1, -1, false, gofpdf.ImageOptions{ImageType: "PNG", ReadDpi: true}, 0, "")

		// ptPdf.SetXY(90, 90)
		ptPdf.Ln(1)
		// fmt.Printf("length of vaccineForWcc %d\n", len(vaccineForWcc))
		ptPdf.SetFontSize(9)
		ptPdf = izTableHeader(ptPdf)
		ptPdf.SetFillColor(255, 255, 255)
		vaccineForWcc := [6]string{" ", " ", " ", " ", " ", " "}
		for _, vaccine := range vaccineForWcc {
			// fmt.Printf("ptName [%s] wccAge [%s] vaccine [%s]\n", apptData.ptName, wccAge, vaccine)
			for i := 0; i < 12; i++ {
				switch i {
				case 0:
					ptPdf.CellFormat(18, 7, vaccine, "1", 0, "", true, 0, "")
				case 1:
					if vaccine != " " {
						ptPdf.CellFormat(18, 7, apptData.apptDate, "1", 0, "", true, 0, "")
					} else {
						ptPdf.CellFormat(18, 7, " ", "1", 0, "", true, 0, "")
					}
				case 2:
					ptPdf.CellFormat(12, 7, " ", "1", 0, "C", true, 0, "")
				case 3:
					ptPdf.CellFormat(12, 7, vacData[strings.ReplaceAll(vaccine, " ", "")].route, "1", 0, "C", true, 0, "")
				case 4:
					ptPdf.CellFormat(18, 7, vacData[strings.ReplaceAll(vaccine, " ", "")].brandName, "1", 0, "C", true, 0, "")
				case 5:
					ptPdf.CellFormat(18, 7, vacData[strings.ReplaceAll(vaccine, " ", "")].manufacturer, "1", 0, "C", true, 0, "")
				case 6:
					ptPdf.CellFormat(18, 7, vacData[strings.ReplaceAll(vaccine, " ", "")].lotNumber, "1", 0, "C", true, 0, "")
				case 7:
					ptPdf.CellFormat(18, 7, vacData[strings.ReplaceAll(vaccine, " ", "")].expiryDt, "1", 0, "C", true, 0, "")
				case 8:
					ptPdf.CellFormat(16, 7, vacData[strings.ReplaceAll(vaccine, " ", "")].visDt, "1", 0, "C", true, 0, "")
				case 9:
					if vaccine != " " {
						ptPdf.CellFormat(17, 7, apptData.apptDate, "1", 0, "", true, 0, "")
					} else {
						ptPdf.CellFormat(17, 7, " ", "1", 0, "", true, 0, "")
					}
					// ptPdf.CellFormat(17, 7, apptData.apptDate,"1", 0, "", true, 0, "")
				case 10:
					ptPdf.CellFormat(12, 7, "", "1", 0, "", true, 0, "")
				case 11:
					ptPdf.CellFormat(18, 7, "", "1", 1, "", true, 0, "")
				}
			}
		}

		err := savePtPDF(ptPdf, apptData.ptName, apptData.apptTime, apptList[0].apptDate, "izadmin", "EMPTY-")
		if err != nil {
			log.Fatalf("Cannot save Patient PDF: %s|n", err)
		}
	}
	return
}

func createVacRefusalPdf(appointments int, apptList [200]SBData, writeFile *os.File, patientList [200]string) {
	var patientMap map[string]bool

	// initialize map
	patientMap = make(map[string]bool)
	for _, s := range patientList {
		patientMap[s] = true
	}

	change2DOSFolder(appointments, apptList)

	for i := 0; i < appointments; i++ {
		ptPdf := gofpdf.New("P", "mm", "Letter", "")
		apptData := apptList[i]
		// fmt.Printf("Looking at patient %s with appt status %s \n",apptData.ptName,  apptData.apptStatus)
		if strings.Contains(apptData.ptName, "Blocked, Appt") || strings.Contains(apptData.visitReason, "Nurse Visit - Non Billable") {
			continue
		}
		apptStr := fmt.Sprintf("Name:%1.30s DOB: %s DOS: %s", apptData.ptName, apptData.dob, apptData.apptDate)
		ptPdf.AddPage()
		ptPdf.SetFont("Times", "", 12)
		ptPdf.SetFontSize(14)
		ptPdf.CellFormat(40, 7, apptStr, "0", 0, "L", false, 0, "")
		ptPdf.Ln(1)
		fileName := fmt.Sprintf("../resource/vac_refusal.png")
		ptPdf.ImageOptions(fileName, 10, 20, -1, -1, false, gofpdf.ImageOptions{ImageType: "PNG", ReadDpi: true}, 0, "")

		err := savePtPDF(ptPdf, apptData.ptName, apptData.apptTime, apptList[0].apptDate, "vac_refusal", "Consent-")
		if err != nil {
			log.Fatalf("Cannot save Patient PDF: %s|n", err)
		}
	}
	return
}

func createFilteredPdf(pdf *gofpdf.Fpdf, appointments int, filterReason, fromTime, toTime string, apptList [200]SBData) *gofpdf.Fpdf {
	var izAdminPageCnt int

	data := loadCSV(path())
	fmt.Printf("FilterReason = [%s]\n", filterReason)
	for i := 0; i < appointments; i++ {
		apptData := apptList[i]
		// fmt.Printf("Looking at patient %s with appt status %s \n",apptData.ptName,  apptData.apptStatus)
		if strings.Contains(apptData.ptName, "Blocked, Appt") {
			continue
		}
		// fmt.Printf("Pt name %s reason [%s] \n", apptData.ptName, apptData.visitReason)
		if !strings.Contains(apptData.visitReason, filterReason) {
			continue
		}
		// Now check if the appointment is within the time period specified. If yes then include else continue
		if isAppointmentInRange(apptData.apptDate, apptData.apptTime, fromTime, toTime) == false {
			continue
		}
		pdf := newReport(pdf, apptData)
		pdf = header(pdf, data[0])
		pdf = table(pdf, data[1:])
		// pdf = image(pdf)
		pdf = sbTableFront(pdf, apptData)
		// till this point total # of pages for the patient is 2
		totalPtPages := 2

		if apptData.isNP {
			if apptData.ptType == "NP" {
				pdf = initialHxImages(pdf, apptData)
				totalPtPages += 2
			}
		}
		if apptData.isWcc {
			if apptData.wccAge != "" {
				if !apptData.isNP && apptData.apptStatus == "Confirmed" {
					/* add bright futures pages */
					pdf = otherImages(pdf, apptData, bfData[apptData.wccAge].numOfPages, "P", true)
					totalPtPages += bfData[apptData.wccAge].numOfPages
					if (totalPtPages % 2) != 0 {
						pdf.AddPage()
						totalPtPages += 1
					}
					/* add bright futures handout pages */
					pdf = otherImages(pdf, apptData, bfData[apptData.wccAge].numOfHandoutPages, "H", false)
					totalPtPages += bfData[apptData.wccAge].numOfHandoutPages
					if (totalPtPages % 2) != 0 {
						pdf.AddPage()
						totalPtPages += 1
					}
					izAdminPageCnt, pdf = addIzAdmin(pdf, apptData)
					totalPtPages += izAdminPageCnt
				} else {
					/* add bright futures pages */
					pdf = otherImages(pdf, apptData, bfData[apptData.wccAge].numOfPages, "P", true)
					totalPtPages += bfData[apptData.wccAge].numOfPages
					if (totalPtPages % 2) != 0 {
						pdf.AddPage()
						totalPtPages += 1
					}
					/* add bright futures handout pages */
					pdf = otherImages(pdf, apptData, bfData[apptData.wccAge].numOfHandoutPages, "H", false)
					totalPtPages += bfData[apptData.wccAge].numOfHandoutPages
					if (totalPtPages % 2) != 0 {
						pdf.AddPage()
						totalPtPages += 1
					}
					izAdminPageCnt, pdf = addIzAdmin(pdf, apptData)
					totalPtPages += izAdminPageCnt
				}
			}
		}
		if apptData.isNV {
			izAdminPageCnt, pdf = addNurseIzAdmin(pdf, apptData)
			totalPtPages += izAdminPageCnt
		}

		if (totalPtPages % 2) != 0 {
			pdf.AddPage()
		}
	}
	return pdf
}

func newReport(pdf *gofpdf.Fpdf, apptData SBData) *gofpdf.Fpdf {

	// AddPage adds a new page to the document. If a page is already present, the
	// Footer() method is called first to output the footer. Then the page is
	// added, the current position set to the top-left corner according to the left
	// and top margins, and Header() is called to display the header.
	//
	pdf.AddPage()

	// SetFont sets the font used to print character strings. It is mandatory to
	// call this method at least once before printing text or the resulting
	// document will not be valid.
	pdf.SetFont("Times", "", 12)

	// CellFormat prints a rectangular cell with optional borders, background color
	// and character string. The upper-left corner of the cell corresponds to the
	// current position. The text can be aligned or centered. After the call, the
	// current position moves to the right or to the next line. It is possible to
	// put a link on the text.
	//

	// Cell is a simpler version of CellFormat with no fill, border, links or
	// special alignment. The Cell_strikeout() example demonstrates this method.
	var apptStr string
	if strings.Contains(apptData.visitReason, "E-Consult") {
		apptStr = fmt.Sprintf("Name:%-40s DOB: %s Time: %s  DOS: %s  %s Pt %s", apptData.ptName, apptData.dob, apptData.apptTime, apptData.apptDate, apptData.ptType, "EV")
	} else if strings.Contains(apptData.visitReason, "Sample") {
		apptStr = fmt.Sprintf("Name:%-40s DOB: %-20s Time: %-18s  DOS: %-18s  %s %s", "", "", "", "", "EP/NP", "WCC/OV/NV/EV")
	} else {
		apptStr = fmt.Sprintf("Name:%-40s DOB: %s Time: %s  DOS: %s  %s Pt %s", apptData.ptName, apptData.dob, apptData.apptTime, apptData.apptDate, apptData.ptType, apptData.apptType)

	}
	// fmt.Printf("Appointment String =[%s]\n", apptStr)

	// pdf.CellFormat(15, 7, "Pt Name: Sanjay Jain     DOB: 11-01-23    Appt Time: 11:00am   DOS: 03-23-23  Est Pt WCC ", "0", 0, "", false, 0, "")
	pdf.CellFormat(15, 7, apptStr, "0", 0, "", false, 0, "")

	// pdf.Cell(40, 10, "DOB")

	// Ln performs a line break. The current abscissa goes back to the left margin
	// and the ordinate increases by the amount passed in parameter. A negative
	// value of h indicates the height of the last printed cell.
	//
	pdf.Ln(10)

	return pdf
}

func header(pdf *gofpdf.Fpdf, hdr []string) *gofpdf.Fpdf {
	pdf.SetFont("Times", "", 14)
	pdf.SetFillColor(220, 220, 220)

	for _, str := range hdr[1:] {
		// pdf.CellFormat(float64(width), 7, str, "LRB", 0, "", true, 0, "")
		pdf.CellFormat(12, 7, str, "1", 0, "", true, 0, "")
	}
	pdf.Ln(-1)
	return pdf
}

func izTableHeader(pdf *gofpdf.Fpdf) *gofpdf.Fpdf {
	pdf.SetFont("Times", "", 9)
	pdf.SetFillColor(240, 240, 240)

	// header row 1
	pdf.CellFormat(18, 7, "", "LTR", 0, "", true, 0, "")
	pdf.CellFormat(42, 7, "Vaccine Administered", "1", 0, "C", true, 0, "")
	pdf.CellFormat(18, 7, "Brand", "LTR", 0, "CB", true, 0, "")
	pdf.CellFormat(54, 7, "Vaccine", "1", 0, "C", true, 0, "")
	pdf.CellFormat(33, 7, "Vaccine Information", "LTR", 0, "C", true, 0, "")
	pdf.CellFormat(12, 7, "Vaccine", "LTR", 0, "C", true, 0, "")
	pdf.CellFormat(18, 7, "Parent/Patient", "LTR", 1, "C", true, 0, "")

	// header row 2
	pdf.CellFormat(18, 7, "", "LR", 0, "C", true, 0, "")

	pdf.CellFormat(18, 7, "Date", "LRT", 0, "CB", true, 0, "")
	pdf.CellFormat(12, 7, "Site on", "LTR", 0, "CB", true, 0, "")
	pdf.CellFormat(12, 7, "", "LTR", 0, "C", true, 0, "")

	pdf.CellFormat(18, 7, "Name", "LR", 0, "CB", true, 0, "")

	// pdf.CellFormat(54, 7, "", "LBR", 0, "C", true, 0, "")
	pdf.CellFormat(18, 7, "", "LR", 0, "C", true, 0, "")
	pdf.CellFormat(18, 7, "", "LR", 0, "C", true, 0, "")
	pdf.CellFormat(18, 7, "BUD(Covid)/", "LR", 0, "C", true, 0, "")
	pdf.CellFormat(33, 7, "Statements Dates", "LRB", 0, "CT", true, 0, "")
	pdf.CellFormat(12, 7, "Admin", "LR", 0, "C", true, 0, "")
	pdf.CellFormat(18, 7, "Guardian", "LR", 1, "C", true, 0, "")

	// header row 3
	pdf.CellFormat(18, 7, "Vaccine", "LBR", 0, "C", true, 0, "")
	pdf.CellFormat(18, 7, "mm/dd/yyyy", "LBR", 0, "C", true, 0, "")
	pdf.CellFormat(12, 7, "Patient", "LBR", 0, "C", true, 0, "")
	pdf.CellFormat(12, 7, "Route", "LBR", 0, "C", true, 0, "")
	pdf.CellFormat(18, 7, "", "LBR", 0, "C", true, 0, "")
	pdf.CellFormat(18, 7, "Manufacturer", "LBR", 0, "C", true, 0, "")
	pdf.CellFormat(18, 7, "Lot Number", "LBR", 0, "C", true, 0, "")
	pdf.CellFormat(18, 7, "Expiration", "LBR", 0, "C", true, 0, "")
	pdf.CellFormat(16, 7, "Published", "LBR", 0, "C", true, 0, "")
	pdf.CellFormat(17, 7, "Provided", "LBR", 0, "C", true, 0, "")
	pdf.CellFormat(12, 7, "Initial", "LBR", 0, "C", true, 0, "")
	pdf.CellFormat(18, 7, "Initials", "LBR", 1, "C", true, 0, "")

	return pdf
}

func sbTableFront(pdf *gofpdf.Fpdf, apptData SBData) *gofpdf.Fpdf {
	pdf.SetFont("Times", "", 9)
	pdf.SetFillColor(240, 240, 240)

	// header row 1
	pdf.CellFormat(24, 5, "WCC NP", "LTR", 0, "", true, 0, "")
	pdf.CellFormat(24, 5, "WCC EP", "LTR", 0, "", true, 0, "")
	pdf.CellFormat(24, 5, "", "LTR", 0, "", true, 0, "")
	pdf.CellFormat(24, 5, "CPT (Vaccine)", "LTR", 0, "", true, 0, "")
	pdf.CellFormat(24, 5, "Vaccine Name", "LTR", 0, "", true, 0, "")
	pdf.CellFormat(24, 5, "MA Init", "LTR", 0, "", true, 0, "")
	pdf.CellFormat(24, 5, "Added 2 OA", "LTR", 0, "", true, 0, "")
	pdf.CellFormat(24, 5, "IZAdm Signed", "LTR", 1, "", true, 0, "")

	pdf.CellFormat(24, 6, "99381", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(19, 6, "99391", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(29, 6, "NB 1 2 4 6 9m", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "90744", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "HBV (Hep B)", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 1, "", false, 0, "")

	pdf.CellFormat(24, 6, "99382", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(19, 6, "99392", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(29, 6, "12 15 18m 2 2.5 3 4y", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "90680", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "Rotateq", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 1, "", false, 0, "")

	pdf.CellFormat(24, 6, "99383", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(19, 6, "99393", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(29, 6, "5 6 7 8 9 10 11y", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "90698", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "Pentacel", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 1, "", false, 0, "")

	pdf.CellFormat(24, 6, "99384", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(19, 6, "99394", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(29, 6, "12 13 14 15 16 17", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "90677", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "PCV 20", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 1, "", false, 0, "")

	pdf.CellFormat(24, 6, "99385", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(19, 6, "99395", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(29, 6, "18y", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 1, "", false, 0, "")

	pdf.CellFormat(14, 5, "OV NP", "LTR", 0, "", true, 0, "")
	pdf.CellFormat(14, 5, "OV EP", "LTR", 0, "", true, 0, "")
	pdf.CellFormat(44, 5, "", "LTR", 0, "", true, 0, "")
	pdf.CellFormat(24, 5, "90633", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 5, "HAV (Hep A)", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 5, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 5, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 5, "", "LTR", 1, "", false, 0, "")

	pdf.CellFormat(14, 6, "99201", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(14, 6, "99211", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(15, 6, "3074F", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(15, 6, "3075F", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(14, 6, "3077F", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "90707", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "MMR", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 1, "", false, 0, "")

	pdf.CellFormat(14, 6, "99202", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(14, 6, "99212", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(15, 6, "3078F", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(15, 6, "3079F", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(14, 6, "3080F", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "90716", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "VZV", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 1, "", false, 0, "")

	pdf.CellFormat(14, 6, "99203", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(14, 6, "99213", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(15, 6, "Mod 25", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(15, 6, "1159F", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(14, 6, "1160F", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "90700", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "DTAP", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 1, "", false, 0, "")

	pdf.CellFormat(14, 6, "99204", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(14, 6, "99214", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(15, 6, "Mod 25", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(15, 6, "4058F", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(14, 6, "4120F", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "90713", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "IPV", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 1, "", false, 0, "")

	pdf.CellFormat(14, 6, "99205", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(14, 6, "99215", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(15, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(15, 6, "4124F", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(14, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "90715", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "TDAP", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 1, "", false, 0, "")

	pdf.CellFormat(11, 5, "CPT", "LTR", 0, "", true, 0, "")
	pdf.CellFormat(22, 5, "Services", "LTR", 0, "", true, 0, "")
	pdf.CellFormat(11, 5, "CPT", "LTR", 0, "", true, 0, "")
	pdf.CellFormat(28, 5, "", "LTR", 0, "", true, 0, "")
	pdf.CellFormat(24, 5, "90651", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 5, "HPV9 (Gardasil)", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 5, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 5, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 5, "", "LTR", 1, "", false, 0, "")

	pdf.CellFormat(11, 6, "94640", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(22, 6, "Neb Tx", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(11, 6, "69210", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(28, 6, "Cerumen Removal", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "90619", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "Mequadfi (MCV4)", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 1, "", false, 0, "")

	pdf.CellFormat(11, 6, "J7613", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(22, 6, "Albuterol Inhal", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(11, 6, "99173", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(28, 6, "Vision Screen", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "90621", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "Trumenba (Men B)", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 1, "", false, 0, "")

	pdf.CellFormat(11, 6, "94760", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(22, 6, "Ox/ 94761 Ox+", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(11, 6, "92552", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(28, 6, "Audio Screen", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "90691", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "Typhoid", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 1, "", false, 0, "")

	pdf.CellFormat(11, 6, "94664", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(22, 6, "MDI Demo", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(11, 6, "10120", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(28, 6, "Splinter Removal", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "90686", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "Flu Regular > 6m", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 1, "", false, 0, "")

	pdf.CellFormat(11, 6, "87636", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(22, 6, "Lucira Cov AB", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(11, 6, "17250", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(28, 6, "Silver Nitrate", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "90674", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "Flu Egg > 6m", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 1, "", false, 0, "")

	pdf.CellFormat(11, 6, "87428", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(22, 6, "CorDX Cov AB", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(11, 6, "24640", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(28, 6, "Nurse Maid Elb", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "91318", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "Covid 6M-4Y", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 1, "", false, 0, "")

	pdf.CellFormat(11, 6, "87804", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(16, 6, "Rapid Flu", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(17, 6, "96110 X 2", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(28, 6, "MCHAT / SWYC", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "91319", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "Covid 5-12Y", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 1, "", false, 0, "")

	pdf.CellFormat(11, 6, "87880", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(16, 6, "Rapid Strep", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(17, 6, "91627", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(28, 6, "PHQ-9", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "90648", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "Act HIB", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "", "LTR", 1, "", false, 0, "")

	pdf.CellFormat(11, 5, "CPT", "LTR", 0, "", true, 0, "")
	pdf.CellFormat(25, 5, "Others", "LTR", 0, "", true, 0, "")
	pdf.CellFormat(11, 5, "CPT", "LTR", 0, "", true, 0, "")
	pdf.CellFormat(25, 5, "", "LTR", 0, "", true, 0, "")
	pdf.CellFormat(24, 5, "CPT", "LTR", 0, "", true, 0, "")
	pdf.CellFormat(96, 5, "", "LTR", 1, "", true, 0, "")

	pdf.CellFormat(11, 6, "17110", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(25, 6, "Wart Removal", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(11, 6, "12001", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(25, 6, "Glue/Dermabond", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "86580", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(96, 6, "PPD", "LTR", 1, "", false, 0, "")

	pdf.CellFormat(11, 6, "69200", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(25, 6, "Fb Removal Ear", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(11, 6, "99177", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(25, 6, "Spot Vision", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "69209", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(96, 6, "Ear Lavage", "LTR", 1, "", false, 0, "")

	pdf.CellFormat(11, 6, "30300", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(25, 6, "FB Removal Nose", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(11, 6, "15020", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(25, 6, "Burn Dressing", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "J0696", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(96, 6, "Ceftriaxone", "LTR", 1, "", false, 0, "")

	pdf.CellFormat(11, 6, "50630", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(25, 6, "Suture Removal", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(11, 6, "98960", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(25, 6, "Epi Pen Demo", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "J1100", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(96, 6, "Dexamethasone", "LTR", 1, "", false, 0, "")

	pdf.CellFormat(120, 6, "", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(24, 6, "MA", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(48, 6, "Mom Dad Br Sis GP", "LTR", 1, "", false, 0, "")

	pdf.CellFormat(120, 5, "", "LR", 0, "", false, 0, "")
	pdf.CellFormat(15, 5, "CC", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(57, 5, "", "LTR", 1, "", false, 0, "")

	pdf.CellFormat(120, 5, "", "LR", 0, "", false, 0, "")
	pdf.CellFormat(15, 5, "Fever", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(57, 5, " x     Day       Max", "LTR", 1, "", false, 0, "")

	pdf.CellFormat(120, 5, "", "LR", 0, "", false, 0, "")
	pdf.CellFormat(15, 5, "Cough", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(57, 5, " x      Day       Mild/ Mod/ Severe", "LTR", 1, "", false, 0, "")

	pdf.CellFormat(120, 5, "", "LR", 0, "", false, 0, "")
	pdf.CellFormat(15, 5, "RN", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(57, 5, " x      Day       Mild/ Mod/ Severe", "LTR", 1, "", false, 0, "")

	pdf.CellFormat(120, 5, "", "LR", 0, "", false, 0, "")
	pdf.CellFormat(15, 5, "Nausea", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(57, 5, " x      Day       Mild/ Mod/ Severe", "LTR", 1, "", false, 0, "")

	pdf.CellFormat(120, 5, "", "LR", 0, "", false, 0, "")
	pdf.CellFormat(15, 5, "Vomit", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(57, 5, " x      Day       Mild/ Mod/ Severe", "LTR", 1, "", false, 0, "")

	pdf.CellFormat(120, 5, "", "LR", 0, "", false, 0, "")
	pdf.CellFormat(15, 5, "Diarrhea", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(57, 5, " x      Day       Mild/ Mod/ Severe", "LTR", 1, "", false, 0, "")

	pdf.CellFormat(120, 5, "", "LR", 0, "", false, 0, "")
	pdf.CellFormat(15, 5, "Constipatn", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(57, 5, " x      Day       Mild/ Mod/ Severe", "LTR", 1, "", false, 0, "")

	pdf.CellFormat(120, 5, "", "LR", 0, "", false, 0, "")
	pdf.CellFormat(15, 5, "BM", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(57, 5, " x      Day       Mild/ Mod/ Severe", "LTR", 1, "", false, 0, "")

	pdf.CellFormat(120, 5, "", "LR", 0, "", false, 0, "")
	pdf.CellFormat(15, 5, "UOP", "LTR", 0, "", false, 0, "")
	pdf.CellFormat(57, 5, " x      Day       Mild/ Mod/ Severe", "LTR", 1, "", false, 0, "")

	pdf.CellFormat(120, 5, "", "LRB", 0, "", false, 0, "")
	pdf.CellFormat(15, 5, "Appetite", "LTRB", 0, "", false, 0, "")
	pdf.CellFormat(57, 5, " x      Day       Mild/ Mod/ Severe", "LTRB", 1, "", false, 0, "")

	if apptData.apptType == "OV" || apptData.apptType == "EV" {
		// fmt.Printf("Copay is [%s] \n", apptData.copay)
		if apptData.copay == "0" {
			footerStr := apptData.ins + " |  " + apptData.inNtwk + " | Coinsurance   " +
				apptData.coins + " | Pt Balance  $" + apptData.ptBalance + " | Sib Balance  $" + apptData.sibBalance
			pdf.CellFormat(192, 5, footerStr, "LTRB", 1, "CB", false, 0, "")
		} else {
			footerStr := apptData.ins + " |  " + apptData.inNtwk + " | Copay  $" +
				apptData.copay + " | Pt Balance  $" + apptData.ptBalance + " | Sib Balance  $" + apptData.sibBalance
			pdf.CellFormat(192, 5, footerStr, "LTRB", 1, "CB", false, 0, "")
		}
	} else {
		footerStr := apptData.ins + " |  " + apptData.inNtwk + " | Pt Balance  $" + apptData.ptBalance + " | Sib Balance  $" + apptData.sibBalance
		pdf.CellFormat(192, 5, footerStr, "LTRB", 1, "CB", false, 0, "")
	}
	// pdf.CellFormat(192, 5, "", "", 1, "CB", false, 0, "")

	pdf.AddPage()

	pdf.CellFormat(32, 5, "Check In", "LRT", 0, "", false, 0, "")
	pdf.CellFormat(32, 5, "Address", "LRT", 0, "", false, 0, "")
	pdf.CellFormat(32, 5, "Insurance", "LRT", 0, "", false, 0, "")
	pdf.CellFormat(32, 5, "US ID", "LRT", 0, "", false, 0, "")
	pdf.CellFormat(32, 5, "Office Policy", "LRT", 0, "", false, 0, "")
	pdf.CellFormat(32, 5, "PT Ally ", "LRT", 1, "", false, 0, "")
	pdf.CellFormat(32, 5, "Initial", "LR", 0, "", false, 0, "")
	pdf.CellFormat(32, 5, "New/Same", "LR", 0, "", false, 0, "")
	pdf.CellFormat(32, 5, "New/Same", "LR", 0, "", false, 0, "")
	pdf.CellFormat(32, 5, "DL/Passport", "LR", 0, "", false, 0, "")
	pdf.CellFormat(32, 5, "Signed (Cur Yr)", "LR", 0, "", false, 0, "")
	pdf.CellFormat(32, 5, "Setup", "LR", 1, "", false, 0, "")
	pdf.CellFormat(32, 5, "", "LR", 0, "", false, 0, "")
	pdf.CellFormat(32, 5, "", "LR", 0, "", false, 0, "")
	pdf.CellFormat(32, 5, "Ins Card Y/N", "LR", 0, "", false, 0, "")
	pdf.CellFormat(32, 5, "  Y/N ", "LR", 0, "", false, 0, "")
	pdf.CellFormat(32, 5, "  Y/N ", "LR", 0, "", false, 0, "")
	pdf.CellFormat(32, 5, "  Y/N ", "LR", 1, "", false, 0, "")
	pdf.CellFormat(32, 5, "", "LRB", 0, "", false, 0, "")
	pdf.CellFormat(32, 5, "", "LRB", 0, "", false, 0, "")
	pdf.CellFormat(32, 5, "", "LRB", 0, "", false, 0, "")
	pdf.CellFormat(32, 5, "", "LRB", 0, "", false, 0, "")
	pdf.CellFormat(32, 5, "", "LRB", 0, "", false, 0, "")
	pdf.CellFormat(32, 5, "", "LRB", 1, "", false, 0, "")
	pdf.CellFormat(96, 15, "Copay/Co-Ins/PR/VP", "LRT", 0, "LT", false, 0, "")
	pdf.CellFormat(96, 15, "Amount", "LRT", 1, "LT", false, 0, "")
	pdf.CellFormat(192, 55, "MA Notes Cont:", "LRT", 1, "LT", false, 0, "")
	pdf.CellFormat(192, 110, "Physician Notes Cont:", "LRTB", 1, "LT", false, 0, "")

	pdf.CellFormat(96, 5, "3074F - Systolic BP < 130mm", "LRT", 0, "", false, 0, "")
	pdf.CellFormat(96, 5, "3078F - Diastolic BP < 80mm", "LRT", 1, "", false, 0, "")
	pdf.CellFormat(96, 5, "3075F - Systolic BP 130 - 139 mm", "LRT", 0, "", false, 0, "")
	pdf.CellFormat(96, 5, "3079F - Diastolic BP 80 - 89mm ", "LRT", 1, "", false, 0, "")
	pdf.CellFormat(96, 5, "3077F - Systolic BP >= 140mm", "LRT", 0, "", false, 0, "")
	pdf.CellFormat(96, 5, "3080F - Diastolic BP >= 90mm", "LRT", 1, "", false, 0, "")
	pdf.CellFormat(96, 5, "1159F - Medication Documented EMR", "LRT", 0, "", false, 0, "")
	pdf.CellFormat(96, 5, "1160F - Medication Reviewed EMR", "LRT", 1, "", false, 0, "")
	pdf.CellFormat(96, 5, "4058F - Pediatrics Gastroenteritis", "LRT", 0, "", false, 0, "")
	pdf.CellFormat(96, 5, "", "LRT", 1, "", false, 1, "")
	pdf.CellFormat(96, 5, "4120F - Pediatrics Pharyngitis (w/ antibiotic)", "LRT", 0, "", false, 0, "")
	pdf.CellFormat(96, 5, "", "LRT", 1, "", false, 1, "")
	pdf.CellFormat(96, 5, "Pediatrics Pharyngitis (w/o antibiotic)", "LRTB", 0, "", false, 0, "")
	pdf.CellFormat(96, 5, "", "LRTB", 1, "", false, 1, "")

	pdf.CellFormat(192, 6, "Arti Pediatrics, Inc Tax Id: 465719521", "", 1, "CB", false, 0, "")
	pdf.CellFormat(192, 6, "2500 Hospital Dr, Ste 8A, Mountain View, CA 94040 P:4084629261 F:4087015006", "", 1, "CB", false, 0, "")

	return pdf
}

func table(pdf *gofpdf.Fpdf, tbl [][]string) *gofpdf.Fpdf {
	pdf.SetFont("Times", "", 16)
	pdf.SetFillColor(255, 255, 255)

	// align := []string{"L", "C", "L", "R", "R", "R", "R","C","R", "L", "L", "R", "R", "R","R", "L", "L"}
	for _, line := range tbl {

		width, _ := strconv.Atoi(line[0])
		for _, str := range line[1:] {
			// pdf.CellFormat(15, 7, str, "1", 0, align[i], false, 0, "")
			pdf.CellFormat(float64(width), 14, str, "1", 0, "C", false, 0, "")
		}
		pdf.Ln(-1)
	}
	return pdf
}

/*
func image(pdf *gofpdf.Fpdf) *gofpdf.Fpdf {
    pdf.ImageOptions("../resource/sb_image.png", 8, 41, -1, -1, false, gofpdf.ImageOptions{ImageType: "PNG", ReadDpi: true}, 0, "")
    pdf.AddPage()
    pdf.ImageOptions("../resource/sb_back.png", 12, 5, -1, -1, false, gofpdf.ImageOptions{ImageType: "PNG", ReadDpi: true}, 0, "")
    return pdf
}
*/

func otherImages(pdf *gofpdf.Fpdf, apptData SBData, numOfPages int, prefix string, appt bool) *gofpdf.Fpdf {
	pdf.SetFontSize(14)
	for i := 0; i < numOfPages; i++ {
		pdf.AddPage()
		if appt {
			apptStr := fmt.Sprintf("Name:%1.30s DOB: %s DOS: %s", apptData.ptName, apptData.dob, apptData.apptDate)

			pdf.CellFormat(15, 7, apptStr, "0", 0, "", false, 0, "")
			pdf.Ln(1)
		}

		fileName := fmt.Sprintf("../BF/%s%s%d.png", apptData.wccAge, prefix, i+1)
		// fmt.Printf("otherImages  wccAge = [%s] for [%s] and fileName [%s] \n", apptData.wccAge, apptData.ptName, fileName)
		pdf.ImageOptions(fileName, 10, 20, -1, -1, false, gofpdf.ImageOptions{ImageType: "PNG", ReadDpi: true}, 0, "")
	}
	return pdf
}

func addIzAdmin(pdf *gofpdf.Fpdf, apptData SBData) (int, *gofpdf.Fpdf) {
	wccAge := apptData.wccAge
	vaccineString := immData[wccAge]
	// fmt.Printf("wccAge [%s] vaccineString [%s] for [%s]\n", wccAge, vaccineString, apptData.ptName)
	if len(vaccineString) == 0 {
		return 0, pdf
	}

	pdf.AddPage()
	pdf.SetFontSize(14)
	pdf.CellFormat(192, 8, "VACCINE ADMINISTRATION RECORD", "", 1, "CB", true, 0, "")
	pdf.CellFormat(192, 8, "Arti Pediatrics, Inc", "", 1, "CB", true, 0, "")
	pdf.CellFormat(192, 8, "2500 Hospital Dr, Suite 8A", "", 1, "CB", true, 0, "")
	pdf.CellFormat(192, 8, "Mountain VIew, CA 94040", "", 1, "CB", true, 0, "")
	pdf.CellFormat(192, 8, "408-462-9261(0) / 408-701-5006 (Fax)", "", 1, "CB", true, 0, "")

	// pdf.SetFontSize(12)
	// apptStr := fmt.Sprintf("Name:%1.30s DOB: %s DOS: %s", apptData.ptName, apptData.dob, apptData.apptDate)
	// pdf.CellFormat(15, 7, apptStr, "0", 0, "", false, 0, "")
	// pdf.Ln(1)
	pdf.CellFormat(192, 7, "", "", 1, "", false, 0, "")
	pdf.Ln(1)
	pdf.SetFontSize(12)
	apptStr := fmt.Sprintf("Patient Name    : %s ", apptData.ptName)
	pdf.CellFormat(120, 7, apptStr, "B", 1, "", false, 0, "")
	apptStr = fmt.Sprintf("Birth Date       : %s ", apptData.dob)
	pdf.CellFormat(120, 7, apptStr, "B", 1, "", false, 0, "")
	apptStr = fmt.Sprintf("Date of Service : %s ", apptData.apptDate)
	pdf.CellFormat(120, 7, apptStr, "B", 1, "", false, 0, "")
	pdf.SetFontSize(9)
	pdf.CellFormat(192, 7, "A copy of the appropriate Centers for Disease Control and Prevention Vaccine Information Statement was provided to me. By signing below, I agree that", "", 1, "", false, 0, "")
	// pdf.CellFormat(192, 7, "By signing below, I agree that", "", 1, "", false, 0, "")
	pdf.CellFormat(192, 7, "* I have read or had explained to me the information about this disease and the vaccine.", "", 1, "B", false, 0, "")
	pdf.CellFormat(192, 7, "* I had an opportunity to ask questions, and those questions were answered satisfactorily.", "", 1, "B", false, 0, "")
	pdf.CellFormat(192, 7, "* I believe that I understand the benefits and risks of the vaccine.", "", 1, "B", false, 0, "")
	pdf.CellFormat(192, 7, "* I ask that the vaccine be given to me or the person named above (for whom I am authorized to make the request).", "", 1, "B", false, 0, "")
	pdf.CellFormat(192, 7, "Every time I initial the \"Parent/Guardian/Patient Initials box, I agree that all of these actions have occured for the vaccine listed in that row.", "", 1, "", false, 0, "")
	pdf.CellFormat(192, 12, "", "", 1, "", false, 0, "")
	pdf.Ln(1)
	pdf.CellFormat(120, 7, "Parent/Guardian/Patient Signature                                                Date", "T", 1, "", false, 0, "")
	// fileName := fmt.Sprintf("../BF/IZAdminHeader.png")
	// pdf.ImageOptions(fileName, 10, 20, -1, -1, false, gofpdf.ImageOptions{ImageType: "PNG", ReadDpi: true}, 0, "")

	vaccineForWcc := strings.Split(vaccineString, ",")
	// fmt.Printf("appStr [%s] wccAge [%s] vaccineString [%s] %s\n", apptStr, wccAge, vaccineString, vaccineForWcc)
	// pdf.SetXY(150, 90)
	pdf.Ln(1)
	// fmt.Printf("length of vaccineForWcc %d\n", len(vaccineForWcc))
	pdf.SetFontSize(9)
	pdf = izTableHeader(pdf)
	pdf.SetFillColor(255, 255, 255)
	for _, vaccine := range vaccineForWcc {
		// fmt.Printf("ptName [%s] wccAge [%s] vaccine [%s]\n", apptData.ptName, wccAge, vaccine)
		for i := 0; i < 12; i++ {
			switch i {
			case 0:
				pdf.CellFormat(18, 7, vaccine, "1", 0, "", true, 0, "")
			case 1:
				// fmt.Printf("ptName [%s] wccAge [%s] vaccine [%s] apptDate [%s]\n", apptData.ptName, wccAge, vaccine, apptData.apptDate)
				pdf.CellFormat(18, 7, apptData.apptDate, "1", 0, "", true, 0, "")
			case 2:
				if wccAge != "4Y" {
					pdf.CellFormat(12, 7, vacData[strings.ReplaceAll(vaccine, " ", "")].site, "1", 0, "C", true, 0, "")
				} else {
					pdf.CellFormat(12, 7, " ", "1", 0, "C", true, 0, "")
				}
			case 3:
				pdf.CellFormat(12, 7, vacData[strings.ReplaceAll(vaccine, " ", "")].route, "1", 0, "C", true, 0, "")
			case 4:
				pdf.CellFormat(18, 7, vacData[strings.ReplaceAll(vaccine, " ", "")].brandName, "1", 0, "C", true, 0, "")
			case 5:
				pdf.CellFormat(18, 7, vacData[strings.ReplaceAll(vaccine, " ", "")].manufacturer, "1", 0, "C", true, 0, "")
			case 6:
				pdf.CellFormat(18, 7, vacData[strings.ReplaceAll(vaccine, " ", "")].lotNumber, "1", 0, "C", true, 0, "")
			case 7:
				pdf.CellFormat(18, 7, vacData[strings.ReplaceAll(vaccine, " ", "")].expiryDt, "1", 0, "C", true, 0, "")
			case 8:
				pdf.CellFormat(16, 7, vacData[strings.ReplaceAll(vaccine, " ", "")].visDt, "1", 0, "C", true, 0, "")
			case 9:
				pdf.CellFormat(17, 7, apptData.apptDate, "1", 0, "", true, 0, "")
			case 10:
				pdf.CellFormat(12, 7, "", "1", 0, "", true, 0, "")
			case 11:
				pdf.CellFormat(18, 7, "", "1", 1, "", true, 0, "")
			}
		}
	}

	for i := 0; i < 5; i++ {
		for i := 0; i < 12; i++ {
			switch i {
			case 0:
				pdf.CellFormat(18, 7, "", "1", 0, "", true, 0, "")
			case 1:
				pdf.CellFormat(18, 7, "", "1", 0, "", true, 0, "")
			case 2:
				pdf.CellFormat(12, 7, "", "1", 0, "", true, 0, "")
			case 3:
				pdf.CellFormat(12, 7, "", "1", 0, "", true, 0, "")
			case 4:
				pdf.CellFormat(18, 7, "", "1", 0, "", true, 0, "")
			case 5:
				pdf.CellFormat(18, 7, "", "1", 0, "", true, 0, "")
			case 6:
				pdf.CellFormat(18, 7, "", "1", 0, "", true, 0, "")
			case 7:
				pdf.CellFormat(18, 7, "", "1", 0, "", true, 0, "")
			case 8:
				pdf.CellFormat(16, 7, "", "1", 0, "", true, 0, "")
			case 9:
				pdf.CellFormat(17, 7, "", "1", 0, "", true, 0, "")
			case 10:
				pdf.CellFormat(12, 7, "", "1", 0, "", true, 0, "")
			case 11:
				pdf.CellFormat(18, 7, "", "1", 1, "", true, 0, "")
			}
		}
	}
	pdf.AddPage()
	return 2, pdf
}

func addAudio(pdf *gofpdf.Fpdf, apptData SBData) (int, *gofpdf.Fpdf) {
	wccAge := strings.ToUpper(apptData.wccAge)

	switch wccAge {
	case "2-5D", "1M", "2M", "4M", "6M", "9M", "12M", "15M", "18M", "2Y", "2.5Y", "3Y":
		return 0, pdf
	}
	pdf.AddPage()
	pdf.SetFontSize(14)
	apptStr := fmt.Sprintf("Name:%1.30s DOB: %s DOS: %s", apptData.ptName, apptData.dob, apptData.apptDate)
	pdf.CellFormat(15, 7, apptStr, "B", 1, "", false, 0, "")
	pdf.Ln(1)
	pdf.Ln(1)
	fileName := fmt.Sprintf("../BF/Audio.png")
	pdf.ImageOptions(fileName, 10, 120, -1, -1, false, gofpdf.ImageOptions{ImageType: "PNG", ReadDpi: true}, 0, "")

	pdf.CellFormat(192, 7, "I authorize staff of Arti Pediatrics to perform audio screening on my child.", "", 1, "", false, 0, "")
	pdf.CellFormat(192, 7, "By signing below, ", "", 1, "", false, 0, "")
	pdf.CellFormat(192, 7, "* I agree that I will be financially liable for the charges in case this service is not covered by insurance", "", 1, "", false, 0, "")
	pdf.CellFormat(192, 12, "", "", 1, "", false, 0, "")
	pdf.Ln(1)
	pdf.Ln(1)
	pdf.Ln(1)
	pdf.CellFormat(120, 7, "Parent/Guardian/Patient Signature                                                Date", "T", 1, "", false, 0, "")
	pdf.CellFormat(192, 12, "", "", 1, "", false, 0, "")
	pdf.Ln(1)
	pdf.CellFormat(192, 7, "* I parent of above mentioned child defer audio screening and I have no hearing concerns for my child", "", 1, "B", false, 0, "")
	pdf.CellFormat(192, 12, "", "", 1, "", false, 0, "")
	pdf.Ln(1)
	pdf.Ln(1)
	pdf.Ln(1)
	pdf.CellFormat(120, 7, "Parent/Guardian/Patient Signature                                                Date", "T", 1, "", false, 0, "")
	pdf.AddPage()

	return 2, pdf
}

func addTB(pdf *gofpdf.Fpdf, apptData SBData) (int, *gofpdf.Fpdf) {
	// fmt.Print("adding TB")
	pdf.AddPage()
	pdf.SetFontSize(14)
	apptStr := fmt.Sprintf("Name:%1.30s DOB: %s DOS: %s", apptData.ptName, apptData.dob, apptData.apptDate)
	fileName := fmt.Sprintf("../BF/TB/tbrisk.png")
	pdf.ImageOptions(fileName, 0, 0, -1, -1, false, gofpdf.ImageOptions{ImageType: "PNG", ReadDpi: true}, 0, "")
	pdf.AddPage()

	pdf.AddPage()
	pdf.SetFontSize(14)
	pdf.CellFormat(15, 7, apptStr, "B", 1, "", false, 0, "")
	fileName = fmt.Sprintf("../BF/TB/tbconsent.png")
	pdf.ImageOptions(fileName, 0, 20, -1, -1, false, gofpdf.ImageOptions{ImageType: "PNG", ReadDpi: true}, 0, "")
	pdf.AddPage()

	pdf.AddPage()
	pdf.SetFontSize(14)
	pdf.CellFormat(15, 7, apptStr, "B", 1, "", false, 0, "")
	fileName = fmt.Sprintf("../BF/TB/tbschool.png")
	pdf.ImageOptions(fileName, 0, 20, -1, -1, false, gofpdf.ImageOptions{ImageType: "PNG", ReadDpi: true}, 0, "")
	pdf.AddPage()

	return 2, pdf
}

func initialHxImages(pdf *gofpdf.Fpdf, apptData SBData) *gofpdf.Fpdf {
	for i := 0; i < 2; i++ {
		pdf.AddPage()
		apptStr := fmt.Sprintf("Name: %s    DOB: %s  DOS: %s ", apptData.ptName, apptData.dob, apptData.apptDate)

		pdf.CellFormat(15, 7, apptStr, "0", 0, "", false, 0, "")
		pdf.Ln(1)

		fileName := fmt.Sprintf("../BF/%sP%d.png", "InitialHx", i+1)
		pdf.ImageOptions(fileName, 10, 20, -1, -1, false, gofpdf.ImageOptions{ImageType: "PNG", ReadDpi: true}, 0, "")
	}
	return pdf
}

func savePDF(pdf *gofpdf.Fpdf, date string, postfix string) error {
	fileName := "Superbills " + strings.Replace(date, "/", "-", 2) + "-" + postfix + ".pdf"
	_, err := os.Stat(fileName)
	if err == nil {
		fmt.Printf("File already exists\n")
		curTime := time.Now()
		hr, min, _ := curTime.Clock()
		hrMinS := fmt.Sprintf("%d-%d", hr, min)
		fileName = "Superbills " + strings.Replace(date, "/", "-", 2) + "-" + postfix + "-" + hrMinS + ".pdf"
	}
	return pdf.OutputFileAndClose(fileName)
}

func savePtPDF(pdf *gofpdf.Fpdf, ptName string, apptTime string, date string, postfix string, prefix string) error {
	fileName := prefix + ptName + "-" + postfix + ".pdf"
	_ = os.Chdir(strings.ReplaceAll(strings.TrimSpace(apptTime), ":", "-") + "-" +
		strings.ReplaceAll(strings.ReplaceAll(strings.TrimSpace(ptName), ",", "-"), " ", "-"))
	_, err := os.Stat(fileName)
	if err == nil {
		// fmt.Printf("File already exists\n")
		curTime := time.Now()
		hr, min, _ := curTime.Clock()
		hrMinS := fmt.Sprintf("%d-%d", hr, min)
		fileName = prefix + ptName + "-" + postfix + "-" + hrMinS + ".pdf"
	}
	err = pdf.OutputFileAndClose(fileName)
	_ = os.Chdir("..")
	return err
}

func loadCSV(path string) [][]string {
	f, err := os.Open(path)
	if err != nil {
		log.Fatalf("Cannot open '%s': %s\n", path, err.Error())
	}
	defer f.Close()
	r := csv.NewReader(f)
	rows, err := r.ReadAll()
	if err != nil {
		log.Fatalln("Cannot read CSV data:", err.Error())
	}
	return rows
}

func path() string {
	return "../resource/Book2.csv"
}

func pathAppt() string {
	if len(os.Args) < 3 {
		return os.Args[1]
	}
	switch len(os.Args) {
	case 2:
		return os.Args[1]
	case 5:
		return os.Args[4]
	}
	return os.Args[len(os.Args)-1]
}

func getCellValue(sheet *xls.Sheet, row int, col string) (string, error) {
	// Get all the rows in the Sheet1.
	colMap := make(map[string]int)
	colMap["A"] = 0
	colMap["B"] = 1
	colMap["C"] = 2
	colMap["D"] = 3
	colMap["E"] = 4
	colMap["F"] = 5
	colMap["G"] = 6
	colMap["H"] = 7
	colMap["I"] = 8
	colMap["J"] = 9
	colMap["K"] = 10
	colMap["L"] = 11
	colMap["M"] = 12
	colMap["N"] = 13
	colMap["O"] = 14
	colMap["P"] = 15
	colMap["Q"] = 16
	colMap["R"] = 17
	colMap["S"] = 18
	colMap["T"] = 19
	colMap["U"] = 20
	colMap["V"] = 21
	colMap["W"] = 22
	colMap["X"] = 23
	colMap["Y"] = 24
	colMap["Z"] = 25

	if row, err := sheet.GetRow(row - 1); err == nil {
		if cell, err := row.GetCol(colMap[col]); err == nil {
			cellStr := cell.GetString()
			return cellStr, err
		}
	}
	return "", nil
}

func loadApptCSV(path string) (int, [200]SBData) {

	workbook, err := xls.OpenFile(path)
	if err != nil {
		fmt.Println(err)
		log.Panic(err.Error())
	}

	sheet, err := workbook.GetSheet(0)
	if err != nil {
		fmt.Println(err)
		panic(err.Error)
	}

	// Get value from cell by given worksheet name and cell reference.
	// func getCellValue (sheet *xls.Sheet, row int,  col string ) string

	a3Str, err := getCellValue(sheet, 3, "A")
	splitStr := strings.Split(a3Str, " ")
	dateOfSvc := splitStr[6]
	if err != nil {
		fmt.Println(err)
		panic(err)
	}

	rows := sheet.GetNumberRows()

	var sbDailyData [200]SBData
	var dailyAppt int = 0

	var sbRow SBData
	for rowN := 0; rowN < rows; rowN++ {
		if rowN < 4 {
			continue
		} else {

			validAppt, _ := getCellValue(sheet, rowN+1, "B")
			apptType, _ := getCellValue(sheet, rowN+1, "K")
			if len(validAppt) == 0 || validAppt == "Total Appointments:" || apptType == "Nurse Visit - Non Billable" {
				continue
			} else {
				sbRow.isNP = false
				sbRow.isNV = false
				sbRow.isWcc = false
				sbRow.apptTime, _ = getCellValue(sheet, rowN+1, "A")
				sbRow.ptID, _ = getCellValue(sheet, rowN+1, "D")
				sbRow.ptName, _ = getCellValue(sheet, rowN+1, "E")
				sbRow.dob, _ = getCellValue(sheet, rowN+1, "H")
				sbRow.apptStatus, _ = getCellValue(sheet, rowN+1, "L")
				sbRow.visitReason, _ = getCellValue(sheet, rowN+1, "K")
				if strings.Contains(sbRow.visitReason, "Nurse") {
					sbRow.apptType = "NV"
					sbRow.isWcc = false
					sbRow.isNV = true
					sbRow.wccAge = getWellChildAge(sbRow.visitReason)
				} else if strings.Contains(sbRow.visitReason, "New") || strings.Contains(sbRow.visitReason, "NP") {
					sbRow.ptType = "NP"
					sbRow.isNP = true
					if strings.Contains(sbRow.visitReason, "Well Child") {
						sbRow.apptType = "WCC"
						sbRow.isWcc = true
						sbRow.wccAge = getWellChildAge(sbRow.visitReason)
					} else {
						sbRow.apptType = "OV"
						sbRow.isWcc = false
						sbRow.wccAge = getWellChildAge(sbRow.visitReason)
					}
				} else if strings.Contains(sbRow.visitReason, "Established") {
					sbRow.ptType = "EP"
					if strings.Contains(sbRow.visitReason, "Well Child") {
						sbRow.apptType = "WCC"
						sbRow.isWcc = true
						sbRow.wccAge = getWellChildAge(sbRow.visitReason)
					} else {
						sbRow.apptType = "OV"
						sbRow.isWcc = false
						sbRow.wccAge = getWellChildAge(sbRow.visitReason)
					}
				} else {
					sbRow.apptType = "OV"
					sbRow.isWcc = false
					sbRow.wccAge = getWellChildAge(sbRow.visitReason)
				}
				sbRow.apptDate = dateOfSvc
				sbDailyData[dailyAppt] = sbRow
				dailyAppt++
			}
		}
	}
	fmt.Printf("Total Appt = %d\n", dailyAppt)

	return dailyAppt, sbDailyData
}

func loadEligibilityReport(path string, dos string, appointments int, apptList [200]SBData) [200]SBData {

	f, err := excelize.OpenFile(path)
	if err != nil {
		fmt.Println(err)
		// panic(err)
		return apptList
	}
	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	// Get value from cell by given worksheet name and cell reference.
	// Get all the rows in the Sheet1.
	colMap := make(map[int]string)
	colMap[0] = "A"
	colMap[1] = "B"
	colMap[2] = "C"
	colMap[3] = "D"
	colMap[4] = "E"
	colMap[5] = "F"
	colMap[6] = "G"
	colMap[7] = "H"
	colMap[8] = "I"
	colMap[9] = "J"
	colMap[10] = "K"
	colMap[11] = "L"
	colMap[12] = "M"
	colMap[13] = "N"
	colMap[14] = "O"
	colMap[15] = "P"
	colMap[16] = "Q"
	colMap[17] = "R"
	colMap[18] = "S"
	colMap[19] = "T"
	colMap[20] = "U"
	colMap[21] = "V"
	colMap[22] = "W"
	colMap[23] = "X"
	colMap[24] = "Y"
	colMap[25] = "Z"
	colMap[26] = "AA"
	rows, err := f.GetRows(dos)
	if err != nil {
		fmt.Println(err)
		// return []
		panic(err)
	}

	for rowN, row := range rows {

		if rowN < 3 {
			continue
		} else {
			cellValue := "B" + strconv.Itoa(rowN+1)
			ptID, _ := f.GetCellValue(dos, cellValue)
			index := 0
			for index < appointments {
				if strings.Compare(apptList[index].ptID, ptID) == 0 {
					break
				} else {
					index++
				}
			}
			for colN, colCell := range row {
				switch colMap[colN] {
				case "N":
					apptList[index].ins = colCell
				case "F":
					apptList[index].inNtwk = colCell
				case "H":
					apptList[index].copay = colCell
					// fmt.Printf("Copay is [%s] [%s]\n", colCell, apptList[index].copay)
				case "I":
					apptList[index].coins = colCell
				case "J":
					apptList[index].ptBalance = colCell
				case "K":
					apptList[index].sibBalance = colCell
				case "E":
					apptList[index].insCard = colCell
				}
			}
		}
	}
	return apptList
}

func loadVaccineData() int {
	f, err := excelize.OpenFile("./resource/VaccineInfo.xlsx")
	if err != nil {
		fmt.Println(err)
		panic(err)
	}
	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	// Get value from cell by given worksheet name and cell reference.
	// Get all the rows in the Sheet1.
	colMap := make(map[int]string)
	colMap[0] = "A"
	colMap[1] = "B"
	colMap[2] = "C"
	colMap[3] = "D"
	colMap[4] = "E"
	colMap[5] = "F"
	colMap[6] = "G"
	colMap[7] = "H"
	colMap[8] = "I"
	colMap[9] = "J"
	colMap[10] = "K"
	colMap[11] = "L"
	colMap[12] = "M"
	colMap[13] = "N"
	colMap[14] = "O"
	colMap[15] = "P"
	colMap[16] = "Q"
	colMap[17] = "R"
	colMap[18] = "S"
	colMap[19] = "T"
	colMap[20] = "U"
	colMap[21] = "V"
	colMap[22] = "W"
	colMap[23] = "X"
	colMap[24] = "Y"
	colMap[25] = "Z"
	colMap[26] = "AA"
	rows, err := f.GetRows("Sheet1")
	if err != nil {
		fmt.Println(err)
		// return []
		panic(err)
	}

	var vacDetails VaccineDetails

	for rowN, row := range rows {

		if rowN < 1 {
			continue
		} else {
			cellValue := "A" + strconv.Itoa(rowN+1)
			vacStr, _ := f.GetCellValue("Sheet1", cellValue)
			for colN, colCell := range row {
				switch colMap[colN] {
				case "B":
					vacDetails.lotNumber = colCell
				case "C":
					vacDetails.qty = colCell
				case "D":
					vacDetails.expiryDt = colCell
				case "E":
					vacDetails.visDt = colCell
				case "F":
					vacDetails.site = colCell
				case "G":
					vacDetails.route = colCell
				case "H":
					vacDetails.brandName = colCell
				case "I":
					vacDetails.manufacturer = colCell
				}
			}
			vacData[vacStr] = vacDetails
		}
		// fmt.Println()
	}
	fmt.Printf("Total Vaccine Keys = %d\n", len(vacData))

	return len(vacData)
}

func getWellChildAge(wellChildStr string) string {
	// fmt.Printf("Find  [%s] to [%s]\n", wellChildStr, strings.ReplaceAll(wellChildStr, " ", ""))
	keys := make([]string, 0, len(bfData))
	for k, _ := range bfData {
		keys = append(keys, k)
	}
	sort.Strings(keys)
	wellChildStr1 := strings.ReplaceAll(wellChildStr, " ", "")
	for _, k := range keys {
		if strings.Contains(strings.ToUpper(wellChildStr1), k) {
			// fmt.Printf("found %s for %s\n", k, wellChildStr)
			return k
		}
	}
	return ""
}

func addNurseIzAdmin(pdf *gofpdf.Fpdf, apptData SBData) (int, *gofpdf.Fpdf) {
	vaccineString := apptData.visitReason

	// fmt.Printf("ptName = [%s] vaccineString = [%s] \n", apptData.ptName, vaccineString)
	if len(vaccineString) == 0 {
		return 0, pdf
	}

	var vaccineForWcc map[string]string
	vaccineForWcc = make(map[string]string)
	vaccineString = strings.Replace(vaccineString, "#", " ", -1)
	vaccineString = strings.Replace(vaccineString, ",", " ", -1)
	vaccineString = strings.Replace(vaccineString, ".", " ", -1)

	scanner := bufio.NewScanner(strings.NewReader(vaccineString))
	scanner.Split(bufio.ScanWords)

	for scanner.Scan() {
		// fmt.Printf("addNurseIzAdmin: tokenStr = [%s] \n", scanner.Text())
		vaccineForWcc[scanner.Text()] = scanner.Text()
	}

	// fmt.Println(vaccineForWcc)

	if !isVaccinesForNV(vaccineString) {
		fmt.Printf("vaccine not found for %s \n", apptData.ptName)
		return 0, pdf
	}

	// fmt.Printf("vaccine found for %s vaccine %s \n", apptData.ptName, vaccineForWcc)

	/*
	   pdf.AddPage()
	   apptStr := fmt.Sprintf("Name:%1.30s DOB: %s DOS: %s", apptData.ptName, apptData.dob, apptData.apptDate)
	   pdf.CellFormat(15, 7, apptStr, "0", 0, "", false, 0, "")
	   pdf.Ln(1)
	   fileName := fmt.Sprintf("../BF/IZAdminHeader.png")
	   pdf.ImageOptions(fileName, 10, 20, -1, -1, false, gofpdf.ImageOptions{ImageType: "PNG", ReadDpi: true}, 0, "")
	*/

	pdf.AddPage()
	pdf.SetFontSize(14)
	pdf.CellFormat(192, 8, "VACCINE ADMINISTRATION RECORD", "", 1, "CT", true, 0, "")
	pdf.CellFormat(192, 8, "Arti Pediatrics, Inc", "", 1, "CT", true, 0, "")
	pdf.CellFormat(192, 8, "2500 Hospital Dr, Suite 8A", "", 1, "CT", true, 0, "")
	pdf.CellFormat(192, 8, "Mountain VIew, CA 94040", "", 1, "CT", true, 0, "")
	pdf.CellFormat(192, 8, "408-462-9261(0) / 408-701-5006 (Fax)", "", 1, "CT", true, 0, "")

	// pdf.SetFontSize(12)
	// apptStr := fmt.Sprintf("Name:%1.30s DOB: %s DOS: %s", apptData.ptName, apptData.dob, apptData.apptDate)
	// pdf.CellFormat(15, 7, apptStr, "0", 0, "", false, 0, "")
	// pdf.Ln(1)
	pdf.CellFormat(192, 7, "", "", 1, "", false, 0, "")
	pdf.Ln(1)
	pdf.SetFontSize(12)
	apptStr := fmt.Sprintf("Patient Name   :%s ", apptData.ptName)
	pdf.CellFormat(120, 7, apptStr, "B", 1, "", false, 0, "")
	apptStr = fmt.Sprintf("Birth Date      :%s ", apptData.dob)
	pdf.CellFormat(120, 7, apptStr, "B", 1, "", false, 0, "")
	apptStr = fmt.Sprintf("Date of Service :%s ", apptData.apptDate)
	pdf.CellFormat(120, 7, apptStr, "B", 1, "", false, 0, "")
	pdf.SetFontSize(9)
	pdf.CellFormat(192, 7, "A copy of the appropriate Centers for Disease Control and Prevention Vaccine Information Statement was provided to me. By signing below, I agree that", "", 1, "", false, 0, "")
	// pdf.CellFormat(192, 7, "By signing below, I agree that", "", 1, "", false, 0, "")
	pdf.CellFormat(192, 7, "* I have read or had explained to me the information about this disease and the vaccine.", "", 1, "B", false, 0, "")
	pdf.CellFormat(192, 7, "* I had an opportunity to ask questions, and those questions were answered satisfactorily.", "", 1, "B", false, 0, "")
	pdf.CellFormat(192, 7, "* I believe that I understand the benefits and risks of the vaccine.", "", 1, "B", false, 0, "")
	pdf.CellFormat(192, 7, "* I ask that the vaccine be given to me or the person named above (for whom I am authorized to make the request).", "", 1, "B", false, 0, "")
	pdf.CellFormat(192, 7, "Every time I initial the \"Parent/Guardian/Patient Initials box, I agree that all of these actions have occured for the vaccine listed in that row.", "", 1, "", false, 0, "")
	pdf.CellFormat(192, 12, "", "", 1, "", false, 0, "")
	pdf.Ln(1)
	pdf.CellFormat(120, 7, "Parent/Guardian/Patient Signature                                                Date", "T", 1, "", false, 0, "")
	// fileName := fmt.Sprintf("../BF/IZAdminHeader.png")
	// pdf.ImageOptions(fileName, 10, 20, -1, -1, false, gofpdf.ImageOptions{ImageType: "PNG", ReadDpi: true}, 0, "")

	// fmt.Printf("appStr [%s] vaccineString [%s] %s\n", apptStr, vaccineString, vaccineForWcc)
	// pdf.SetXY(90, 90)
	pdf.Ln(1)
	// fmt.Printf("length of vaccineForWcc %d\n", len(vaccineForWcc))
	pdf.SetFontSize(9)
	pdf = izTableHeader(pdf)
	pdf.SetFillColor(255, 255, 255)
	covidVac := false
	for _, vaccine := range vaccineForWcc {
		// fmt.Printf("ptName [%s] wccAge [%s] vaccine [%s]\n", apptData.ptName, wccAge, vaccine)
		// fmt.Printf("ptName [%s]  vaccine [%s]\n", apptData.ptName, vaccine)

		vaccineS := getVaccinesForNV(vaccine)
		if vaccineS == "NOT_FOUND" {
			continue
		}
		if vaccineS == "Covid" {
			covidVac = true
		}
		for i := 0; i < 12; i++ {
			switch i {
			case 0:
				pdf.CellFormat(18, 7, vaccineS, "1", 0, "", true, 0, "")
			case 1:
				pdf.CellFormat(18, 7, apptData.apptDate, "1", 0, "", true, 0, "")
			case 2:
				pdf.CellFormat(12, 7, vacData[strings.ReplaceAll(vaccineS, " ", "")].site, "1", 0, "C", true, 0, "")
			case 3:
				pdf.CellFormat(12, 7, vacData[strings.ReplaceAll(vaccineS, " ", "")].route, "1", 0, "C", true, 0, "")
			case 4:
				pdf.CellFormat(18, 7, vacData[strings.ReplaceAll(vaccineS, " ", "")].brandName, "1", 0, "C", true, 0, "")
			case 5:
				pdf.CellFormat(18, 7, vacData[strings.ReplaceAll(vaccineS, " ", "")].manufacturer, "1", 0, "C", true, 0, "")
			case 6:
				pdf.CellFormat(18, 7, vacData[strings.ReplaceAll(vaccineS, " ", "")].lotNumber, "1", 0, "C", true, 0, "")
			case 7:
				pdf.CellFormat(18, 7, vacData[strings.ReplaceAll(vaccineS, " ", "")].expiryDt, "1", 0, "C", true, 0, "")
			case 8:
				pdf.CellFormat(16, 7, vacData[strings.ReplaceAll(vaccineS, " ", "")].visDt, "1", 0, "C", true, 0, "")
			case 9:
				pdf.CellFormat(17, 7, apptData.apptDate, "1", 0, "", true, 0, "")
			case 10:
				pdf.CellFormat(12, 7, "", "1", 0, "", true, 0, "")
			case 11:
				pdf.CellFormat(18, 7, "", "1", 1, "", true, 0, "")
			}
		}
	}

	for i := 0; i < 6; i++ {
		for i := 0; i < 12; i++ {
			switch i {
			case 0:
				pdf.CellFormat(18, 7, "", "1", 0, "", true, 0, "")
			case 1:
				pdf.CellFormat(18, 7, "", "1", 0, "", true, 0, "")
			case 2:
				pdf.CellFormat(12, 7, "", "1", 0, "", true, 0, "")
			case 3:
				pdf.CellFormat(12, 7, "", "1", 0, "", true, 0, "")
			case 4:
				pdf.CellFormat(18, 7, "", "1", 0, "", true, 0, "")
			case 5:
				pdf.CellFormat(18, 7, "", "1", 0, "", true, 0, "")
			case 6:
				pdf.CellFormat(18, 7, "", "1", 0, "", true, 0, "")
			case 7:
				pdf.CellFormat(18, 7, "", "1", 0, "", true, 0, "")
			case 8:
				pdf.CellFormat(16, 7, "", "1", 0, "", true, 0, "")
			case 9:
				pdf.CellFormat(17, 7, "", "1", 0, "", true, 0, "")
			case 10:
				pdf.CellFormat(12, 7, "", "1", 0, "", true, 0, "")
			case 11:
				pdf.CellFormat(18, 7, "", "1", 1, "", true, 0, "")
			}
		}
	}
	if covidVac {
		pdf.AddPage()
		pdf.AddPage()
		pdf.SetFontSize(14)
		apptStr := fmt.Sprintf("Patient Name   :%s Birth Date : %s DOS :%s ", apptData.ptName, apptData.dob, apptData.apptDate)
		pdf.CellFormat(40, 7, apptStr, "0", 0, "L", false, 0, "")
		pdf.Ln(1)
		fileName := fmt.Sprintf("../BF/CovidP1.png")
		pdf.ImageOptions(fileName, 10, 20, -1, -1, false, gofpdf.ImageOptions{ImageType: "PNG", ReadDpi: true}, 0, "")
		pdf.AddPage()
		pdf.AddPage()
		pdf.CellFormat(40, 7, apptStr, "0", 0, "L", false, 0, "")
		pdf.Ln(1)
		fileName = fmt.Sprintf("../BF/CovidP2.png")
		pdf.ImageOptions(fileName, 10, 20, -1, -1, false, gofpdf.ImageOptions{ImageType: "PNG", ReadDpi: true}, 0, "")
		return 4, pdf
	}
	return 1, pdf
}

func isVaccinesForNV(vaccineString string) bool {
	// fmt.Printf("Find  [%s] to [%s]\n", wellChildStr, strings.ReplaceAll(wellChildStr, " ", ""))
	// fmt.Printf("vaccineString [%s] \n", vaccineString)
	keys := make([]string, 0, len(vacData))
	for k, _ := range vacData {
		keys = append(keys, k)
	}
	sort.Strings(keys)

	startIdx := strings.Index(vaccineString, "Nurse Visit ")

	vaccineForWcc := strings.Split(vaccineString[startIdx+len("Nurse Visit ")+1:], " ")
	for _, vaccine := range vaccineForWcc {
		for _, k := range keys {
			// fmt.Printf("key1 is [%s] vaccine1 is [%s]\n", k, vaccine)
			if strings.Compare(strings.TrimSpace((strings.ToUpper(vaccine))), strings.TrimSpace(strings.ToUpper(k))) == 0 {
				// fmt.Printf("returning true from isVaccinesForNV \n")
				return true
			}
		}
	}

	vaccineForWcc2 := strings.Split(vaccineString[startIdx+len("Nurse Visit ")+1:], ",")
	for _, vaccine := range vaccineForWcc2 {
		for _, k := range keys {
			// fmt.Printf("key2 is [%s] vaccine2 is [%s]\n", k, vaccine)
			if strings.Compare(strings.TrimSpace((strings.ToUpper(vaccine))), strings.TrimSpace(strings.ToUpper(k))) == 0 {
				// fmt.Printf("2 returning true from isVaccinesForNV \n")
				return true
			}
		}
	}
	return false
}
func getVaccinesForNV(vaccineString string) string {
	// fmt.Printf("Find  [%s] \n", vaccineString)
	keys := make([]string, 0, len(vacData))
	for k, _ := range vacData {
		keys = append(keys, k)
	}
	sort.Strings(keys)
	vaccineForWcc := strings.Split(vaccineString, " ")
	for _, vaccine := range vaccineForWcc {
		for _, k := range keys {
			if strings.Compare(strings.ToUpper(vaccine), strings.ToUpper(k)) == 0 {
				return k
			}
		}
	}
	return "NOT_FOUND"
}

func isAppointmentInRange(dos, apptTime, fromTime, toTime string) bool {
	// fmt.Printf("Inside isAppointmentInRange [%s] [%s] [%s]\n", apptTime, fromTime, toTime)

	dosSplitArr := strings.Split(dos, "/")
	timeSplitArr := strings.Split(apptTime, ":")
	fromTimeSplitArr := strings.Split(fromTime, ":")
	toTimeSplitArr := strings.Split(toTime, ":")

	var mmDosStr, ddDosStr, normDosStr, normApptTime, hhTimeStr string
	var normFromTime, normToTime string

	if len(dosSplitArr[0]) != 0 {
		mmDosStr = "0" + dosSplitArr[0]
	} else {
		mmDosStr = dosSplitArr[0]
	}

	if len(dosSplitArr[1]) != 0 {
		ddDosStr = "0" + dosSplitArr[1]
	} else {
		ddDosStr = dosSplitArr[1]
	}

	normDosStr = mmDosStr + "/" + ddDosStr + "/" + dosSplitArr[2]

	if len(strings.TrimSpace(timeSplitArr[0])) == 1 {
		hhTimeStr = "0" + strings.TrimSpace(timeSplitArr[0])
	} else {
		hhTimeStr = timeSplitArr[0]
	}

	normApptTime = hhTimeStr + ":" + timeSplitArr[1]

	if len(fromTimeSplitArr[0]) == 1 {
		hhTimeStr = "0" + fromTimeSplitArr[0]
	} else {
		hhTimeStr = fromTimeSplitArr[0]
	}

	normFromTime = hhTimeStr + ":" + fromTimeSplitArr[1]

	if len(toTimeSplitArr[0]) == 1 {
		hhTimeStr = "0" + toTimeSplitArr[0]
	} else {
		hhTimeStr = toTimeSplitArr[0]
	}
	normToTime = hhTimeStr + ":" + toTimeSplitArr[1]

	fmt.Printf("NormDosStr [%s] NormApptTime [%s] NormFromTime [%s] NormToTime [%s] \n", normDosStr, normApptTime, normFromTime, normToTime)

	apptTimeType, _ := time.Parse("01/02/2006 03:04pm", normDosStr+" "+normApptTime)
	fromTimeType, _ := time.Parse("01/02/2006 03:04pm", normDosStr+" "+normFromTime)
	toTimeType, _ := time.Parse("01/02/2006 03:04pm", normDosStr+" "+normToTime)

	if (apptTimeType.Equal(fromTimeType) || apptTimeType.After(fromTimeType)) && (apptTimeType.Equal(toTimeType) || apptTimeType.Before(toTimeType)) {
		fmt.Printf("Appointment to be included [%s] from [%s] and to [%s]\n", normApptTime, normFromTime, normToTime)
		return true
	}
	return false
}

func readCreatedSBsFile(path string) (*os.File, [200]string) {

	// workbook, err := xls.OpenFile("Report_DailyAppointments-380668.xls")
	workbook, err := xls.OpenFile(path)
	if err != nil {
		fmt.Println(err)
		log.Panic(err.Error())
	}

	sheet, err := workbook.GetSheet(0)
	if err != nil {
		fmt.Println(err)
		panic(err.Error)
	}

	var dateOfSvc string
	if r3, err := sheet.GetRow(2); err == nil {
		if cell, err := r3.GetCol(0); err == nil {
			a3Str := cell.GetString()
			splitStr := strings.Split(a3Str, " ")
			dateOfSvc = splitStr[6]
		}
	}

	normDos := strings.Replace(dateOfSvc, "/", "-", -1)
	err = os.Chdir(normDos)
	if err != nil {
		log.Fatal(err)
	}
	f1, err1 := os.OpenFile("superbillApps-"+normDos+".log", os.O_CREATE|os.O_RDWR, 0666)
	if err1 != nil {
		log.Fatal(err)
	}
	f1.Seek(0, os.SEEK_SET)

	fileScanner := bufio.NewScanner(f1)

	var patientList [200]string
	i := 0
	for fileScanner.Scan() {
		patientList[i] = fileScanner.Text()
		// fmt.Println(fileScanner.Text())
		i++
	}
	_ = os.Chdir("..")
	return f1, patientList
}

func createDOSDir(path string) (string, error) {
	workbook, err := xls.OpenFile(path)
	if err != nil {
		fmt.Println(err)
		log.Panic(err.Error())
	}

	sheet, err := workbook.GetSheet(0)
	if err != nil {
		fmt.Println(err)
		panic(err.Error)
	}

	var dateOfSvc string
	if r3, err := sheet.GetRow(2); err == nil {
		if cell, err := r3.GetCol(0); err == nil {
			a3Str := cell.GetString()
			splitStr := strings.Split(a3Str, " ")
			dateOfSvc = splitStr[6]
		}
	}
	normDos := strings.Replace(dateOfSvc, "/", "-", -1)

	err = os.Mkdir(normDos, 0777)
	if err != nil && !os.IsExist(err) {
		log.Fatal(err)
	}
	if os.IsExist(err) {
		fmt.Printf("Cleaning existing DOS directory %s \n", normDos)
		os.RemoveAll(normDos)
		err = os.Mkdir(normDos, 0777)
		if err != nil && !os.IsExist(err) {
			log.Fatal(err)
		}
	}
	return dateOfSvc, err
}
