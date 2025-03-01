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

func extractBranch(campusid string) string {
	if len(campusid) >= 6 {
		return campusid[4:6]
	}
	return ""
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

	if len(errors) > 0 {
		log.Println("Errors found during processing:")
		for _, err := range errors {
			log.Println(" ", err)
		}
	}

	if len(students) == 0 {
		log.Println("No student records found.")
	} else {
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

func parseExcel(filePath string, filterClass string) ([]Student, []string) {
	var students []Student
	var errors []string

	f, err := excelize.OpenFile(filePath)
	if err != nil {
		log.Fatalf("Failed to open file: %v", err)
	}
	defer f.Close()

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

		computedTotal := quiz + midSem + labTest + weeklyLabs + compre
		if computedTotal != total {
			errors = append(errors, fmt.Sprintf("Error: Mismatch for CAMPUSID %s -> Expected %d, Found %d", campusid, computedTotal, total))
		}

		students = append(students, Student{CampusID: campusid, ClassNo: row[indexMap["ClassNo."]], Branch: extractBranch(campusid), Quiz: quiz, MidSem: midSem, LabTest: labTest, WeeklyLabs: weeklyLabs, PreCompre: preCompre, Compre: compre, Total: total})
	}
	return students, errors
}

func generateReport(students []Student, errors []string) Report {
	rankings := make(map[string][]RankingEntry)
	components := []string{"Quiz", "MidSem", "LabTest", "WeeklyLabs", "PreCompre", "Compre", "Total"}

	for _, comp := range components {
		scores := make([]RankingEntry, len(students))
		for i, s := range students {
			var score int
			switch comp {
			case "Quiz":
				score = s.Quiz
			case "MidSem":
				score = s.MidSem
			case "LabTest":
				score = s.LabTest
			case "WeeklyLabs":
				score = s.WeeklyLabs
			case "PreCompre":
				score = s.PreCompre
			case "Compre":
				score = s.Compre
			case "Total":
				score = s.Total
			}
			scores[i] = RankingEntry{CampusID: s.CampusID, Score: score}
		}

		sort.Slice(scores, func(i, j int) bool { return scores[i].Score > scores[j].Score })

		if len(scores) > 0 {
			rankings[comp] = append(rankings[comp], scores[0])
		}
		if len(scores) > 1 {
			rankings[comp] = append(rankings[comp], scores[1])
		}
		if len(scores) > 2 {
			rankings[comp] = append(rankings[comp], scores[2])
		}
	}

	return Report{Rankings: rankings, Errors: errors}
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
	file, err := os.Create("report.json")
	if err != nil {
		log.Fatalf("Error creating report.json: %v", err)
	}
	defer file.Close()

	encoder := json.NewEncoder(file)
	encoder.SetIndent("", "  ")
	err = encoder.Encode(report)
	if err != nil {
		log.Fatalf("Error writing JSON to file: %v", err)
	}

	log.Println("Report successfully exported to report.json")
}
