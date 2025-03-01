/*************  ✨ Codeium Command ⭐  *************/
package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"log"
	"os"
	"sort"
	"strconv"

	"github.com/xuri/excelize/v2"
)

type Student struct {
	/*************  ✨ Codeium Command 🌟  *************/
	CampusID   string
	ClassNo    string
	Branch     string
	Quiz       int
	MidSem     int
	LabTest    int
	WeeklyLabs int
	PreCompre  int
	Compre     int
	Total      int
}

type Report struct {
	Averages       map[string]float64
	BranchAverages map[string]float64
	Rankings       map[string][]RankingEntry
	Errors         []string
}

type RankingEntry struct {
	CampusID string
	Score    int
	Rank     string
}

func main() {
	exportJSON := flag.Bool("export", false, "Export summary report to JSON")
	filterClass := flag.String("class", "", "Filter by ClassNo.")
	flag.Parse()

	if flag.NArg() < 1 {
		log.Fatal("Usage: program <CSF111_202425_01_GradeBook_stripped.xlsx> [--export] [--class=ClassNo.]")
	}
	filePath := flag.Arg(0)

	students, errors := parseExcel(filePath, *filterClass)
	for i := range students {
		students[i].Total = students[i].PreCompre + students[i].Compre
	}
	/*************  ✨ Codeium Command 🌟  *************/
	if len(errors) > 0 {
		log.Println("Errors found during processing:")
		/*************  ✨ Codeium Command 🌟  *************/
		for _, err := range errors {
			log.Println(" ", err)
		}
	}
	if len(students) == 0 {
		log.Println("No student records found.")
	} else {
		// Update CampusID and calculate Branch
		for i := range students {
			students[i].Branch = extractBranch(students[i].CampusID)
		}

		report := generateReport(students, errors)
		printReport(report)

		if *exportJSON {
			exportReport(report)
		}
	}
}

/******  3615111c-df86-4064-99e0-b8678f35c50a  *******/

func parseExcel(filePath string, filterClass string) ([]Student, []string) {
	var students []Student
	var errors []string

	f, err := excelize.OpenFile(filePath)
	if err != nil {
		log.Fatalf("Failed to open file: %v", err)
	}
	defer func() {
		if err := f.Close(); err != nil {
			log.Printf("Failed to close file: %v", err)
		}
	}()

	sheetName := f.GetSheetName(0)
	rows, err := f.GetRows(sheetName)
	if err != nil {
		log.Fatalf("Failed to read rows: %v", err)
	}

	if len(rows) == 0 {
		log.Fatal("The sheet is empty.")
	}

	header := rows[0]
	indexMap := make(map[string]int)
	for i, colName := range header {
		indexMap[colName] = i
	}

	required := []string{"CampusID", "ClassNo.", "Quiz", "MidSem", "LabTest", "WeeklyLabs", "PreCompre", "Compre", "Total"}
	for _, col := range required {
		if _, ok := indexMap[col]; !ok {
			log.Fatalf("Missing required column: %s", col)
		}
	}

	for _, row := range rows[1:] {
		if len(row) == 0 {
			continue
		}

		campusid := row[indexMap["CampusID"]]
		if filterClass != "" && row[indexMap["ClassNo."]] != filterClass {
			continue
		}

		quiz, _ := strconv.Atoi(row[indexMap["Quiz"]])
		midSem, _ := strconv.Atoi(row[indexMap["MidSem"]])
		labTest, _ := strconv.Atoi(row[indexMap["LabTest"]])
		weeklyLabs, _ := strconv.Atoi(row[indexMap["WeeklyLabs"]])
		preCompre, _ := strconv.Atoi(row[indexMap["PreCompre"]])
		compre, _ := strconv.Atoi(row[indexMap["Compre"]])
		total, _ := strconv.Atoi(row[indexMap["Total"]])

		computedTotal := preCompre + compre
		if computedTotal != total {
			errors = append(errors, fmt.Sprintf("Error: Mismatch for CAMPUSID %s -> Expected %d, Found %d", campusid, computedTotal, total))
		}

		students = append(students, Student{
			CampusID:   campusid,
			ClassNo:    row[indexMap["ClassNo."]],
			Branch:     extractBranch(campusid),
			Quiz:       quiz,
			MidSem:     midSem,
			LabTest:    labTest,
			WeeklyLabs: weeklyLabs,
			PreCompre:  preCompre,
			Compre:     compre,
			Total:      total,
		})
	}
	return students, errors
}

func extractBranch(campusid string) string {
	if len(campusid) >= 6 {
		return campusid[4:6]
	}
	return ""
}

func generateReport(students []Student, errors []string) Report {
	var totalQuiz, totalMid, totalLab, totalWeekly, totalPre, totalCompre, totalOverall int
	count := len(students)

	branchTotals := make(map[string]int)
	branchCounts := make(map[string]int)

	type compScore struct {
		CampusID string
		Score    int
	}

	ranking := map[string][]compScore{
		"Quiz":       {},
		"MidSem":     {},
		"LabTest":    {},
		"WeeklyLabs": {},
		"PreCompre":  {},
		"Compre":     {},
		"Total":      {},
	}

	for _, s := range students {
		totalQuiz += s.Quiz
		totalMid += s.MidSem
		totalLab += s.LabTest
		totalWeekly += s.WeeklyLabs
		totalPre += s.PreCompre
		totalCompre += s.Compre
		totalOverall += s.Total

		branchTotals[s.Branch] += s.Total
		branchCounts[s.Branch]++

		ranking["Quiz"] = append(ranking["Quiz"], compScore{s.CampusID, s.Quiz})
		ranking["MidSem"] = append(ranking["MidSem"], compScore{s.CampusID, s.MidSem})
		ranking["LabTest"] = append(ranking["LabTest"], compScore{s.CampusID, s.LabTest})
		ranking["WeeklyLabs"] = append(ranking["WeeklyLabs"], compScore{s.CampusID, s.WeeklyLabs})
		ranking["PreCompre"] = append(ranking["PreCompre"], compScore{s.CampusID, s.PreCompre})
		ranking["Compre"] = append(ranking["Compre"], compScore{s.CampusID, s.Compre})
		ranking["Total"] = append(ranking["Total"], compScore{s.CampusID, s.Total})
	}

	averages := map[string]float64{
		"Quiz":       float64(totalQuiz) / float64(count),
		"MidSem":     float64(totalMid) / float64(count),
		"LabTest":    float64(totalLab) / float64(count),
		"WeeklyLabs": float64(totalWeekly) / float64(count),
		"PreCompre":  float64(totalPre) / float64(count),
		"Compre":     float64(totalCompre) / float64(count),
		"Total":      float64(totalOverall) / float64(count),
	}

	branchAverages := make(map[string]float64)
	for branch, total := range branchTotals {
		branchAverages[branch] = float64(total) / float64(branchCounts[branch])
	}

	rankings := make(map[string][]RankingEntry)
	for comp, scores := range ranking {
		sort.Slice(scores, func(i, j int) bool {
			return scores[i].Score > scores[j].Score
		})

		if len(scores) > 0 {
			rankings[comp] = append(rankings[comp], RankingEntry{scores[0].CampusID, scores[0].Score, "1st"})
		}
		if len(scores) > 1 {
			rankings[comp] = append(rankings[comp], RankingEntry{scores[1].CampusID, scores[1].Score, "2nd"})
		}
		if len(scores) > 2 {
			rankings[comp] = append(rankings[comp], RankingEntry{scores[2].CampusID, scores[2].Score, "3rd"})
		}
	}

	return Report{Averages: averages, BranchAverages: branchAverages, Rankings: rankings, Errors: errors}
}

func printReport(report Report) {
	fmt.Println("Summary Report:")
	fmt.Println("Overall Averages:")
	for comp, avg := range report.Averages {
		fmt.Printf(" %s: %.2f\n", comp, avg)
	}

	fmt.Println("\nBranch-wise Averages (Total Scores):")
	for branch, avg := range report.BranchAverages {
		fmt.Printf(" Branch %s: %.2f\n", branch, avg)
	}

	fmt.Println("\nTop 3 Rankings per Component:")
	for comp, entries := range report.Rankings {
		fmt.Printf(" %s:\n", comp)
		for _, entry := range entries {
			fmt.Printf("  %s - %d (%s)\n", entry.CampusID, entry.Score, entry.Rank)
		}
	}

	if len(report.Errors) > 0 {
		fmt.Println("\nData Validation Errors:")
		for _, errMsg := range report.Errors {
			fmt.Println(" ", errMsg)
		}
	}
}

func exportReport(report Report) {
	data, err := json.MarshalIndent(report, "", " ")
	if err != nil {
		log.Printf("Failed to marshal report to JSON: %v", err)
		return
	}
	fileName := "report.json"
	err = os.WriteFile(fileName, data, 0644)
	if err != nil {
		log.Printf("Failed to write JSON report to file: %v", err)
		return
	}
	fmt.Printf("\nExported report to %s\n", fileName)
}

/******  710a782b-7c8c-41f6-8af8-4af5bfbd87f2  *******/
