package main

import (
	"fmt"
	"sort"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

func computeAverages(rows [][]string) (count int, quizAvg, midSemAvg, labTestAvg, labsAvg, preAvg, compreAvg, totalAvg float64, errors []string) {

	var quizTotal, midSemTotal, labTestTotal, labsTotal, preTotal, compreTotal, grandTotal float64

	// Iterate through the dataset (skipping the header row if necessary).
	for i := 1; i < len(rows); i++ {
		row := rows[i]
		if len(row) < 11 {
			continue
		}

		// actualRow is used for user-friendly row numbering.
		actualRow := i + 1
		id := strings.TrimSpace(row[0])

		quiz, err := strconv.ParseFloat(strings.TrimSpace(row[4]), 64)
		if err != nil {
			fmt.Printf("Row %d error in Quiz: %v\n", actualRow, err)
			continue
		}
		midSem, err := strconv.ParseFloat(strings.TrimSpace(row[5]), 64)
		if err != nil {
			fmt.Printf("Row %d error in Mid-Sem: %v\n", actualRow, err)
			continue
		}
		labTest, err := strconv.ParseFloat(strings.TrimSpace(row[6]), 64)
		if err != nil {
			fmt.Printf("Row %d error in Lab Test: %v\n", actualRow, err)
			continue
		}
		labs, err := strconv.ParseFloat(strings.TrimSpace(row[7]), 64)
		if err != nil {
			fmt.Printf("Row %d error in Weekly Labs: %v\n", actualRow, err)
			continue
		}
		pre, err := strconv.ParseFloat(strings.TrimSpace(row[8]), 64)
		if err != nil {
			fmt.Printf("Row %d error in Pre-Compre: %v\n", actualRow, err)
			continue
		}
		compre, err := strconv.ParseFloat(strings.TrimSpace(row[9]), 64)
		if err != nil {
			fmt.Printf("Row %d error in Compre: %v\n", actualRow, err)
			continue
		}
		total, err := strconv.ParseFloat(strings.TrimSpace(row[10]), 64)
		if err != nil {
			fmt.Printf("Row %d error in Total (300): %v\n", actualRow, err)
			continue
		}

		// Calculate the expected total from individual components.
		calculated := quiz + midSem + labTest + labs + compre
		if calculated != total {
			errors = append(errors, fmt.Sprintf("Row %d (ID %s): calculated %.2f != total %.2f", actualRow, id, calculated, total))
		}

		// Accumulate totals.
		quizTotal += quiz
		midSemTotal += midSem
		labTestTotal += labTest
		labsTotal += labs
		preTotal += pre
		compreTotal += compre
		grandTotal += total

		count++
	}

	// Calculate averages if any valid records were processed.
	if count > 0 {
		quizAvg = quizTotal / float64(count)
		midSemAvg = midSemTotal / float64(count)
		labTestAvg = labTestTotal / float64(count)
		labsAvg = labsTotal / float64(count)
		preAvg = preTotal / float64(count)
		compreAvg = compreTotal / float64(count)
		totalAvg = grandTotal / float64(count)
	}

	return
}

func Top3(rows [][]string, component string) [][]string {
	componentmap := make(map[string]float64)
	componentmap["quiz"] = 4
	componentmap["midSem"] = 5
	componentmap["labTest"] = 6
	componentmap["weeklylabs"] = 7
	componentmap["precompre"] = 8
	componentmap["compre"] = 9
	componentmap["total"] = 10
	type rowScore struct {
		row   []string
		score float64
	}
	var data []rowScore
	for i := 1; i < len(rows); i++ {
		s := strings.TrimSpace(rows[i][int(componentmap[component])])
		score, err := strconv.ParseFloat(s, 64)
		if err != nil {
			continue
		}
		data = append(data, rowScore{row: rows[i], score: score})
	}

	sort.Slice(data, func(i, j int) bool { return data[i].score > data[j].score })

	var topRows [][]string
	for i := 0; i < 3; i++ {
		topRows = append(topRows, data[i].row)
	}
	return topRows
}

func main() {
	// Expect the file path as the first command-line argument.
	//if len(os.Args) < 2 {
	//	fmt.Println("Usage: go run main.go <file_path>")
	//	return
	//}
	//filePath := os.Args[1]

	// Open the Excel file.
	f, err := excelize.OpenFile("CSgradebook.xlsx")
	if err != nil {
		fmt.Println("Error opening file:", err)
		return
	}
	defer f.Close()

	sheet := f.GetSheetName(0)
	if sheet == "" {
		fmt.Println("No sheet found")
		return
	}

	rows, err := f.GetRows(sheet)
	if err != nil {
		fmt.Println("Error reading rows:", err)
		return
	}

	if len(rows) < 2 {
		fmt.Println("Not enough rows")
		return
	}

	count, quizAvg, midSemAvg, labTestAvg, labsAvg, preAvg, compreAvg, totalAvg, errors := computeAverages(rows)
	fmt.Println("Overall Records:", count)
	if count > 0 {
		fmt.Println("Overall Averages:")
		fmt.Printf("Quiz: %.2f\n", quizAvg)
		fmt.Printf("Mid-Sem: %.2f\n", midSemAvg)
		fmt.Printf("Lab Test: %.2f\n", labTestAvg)
		fmt.Printf("Weekly Labs: %.2f\n", labsAvg)
		fmt.Printf("Pre-Compre: %.2f\n", preAvg)
		fmt.Printf("Compre: %.2f\n", compreAvg)
		fmt.Printf("Total: %.2f\n", totalAvg)
	}
	if len(errors) > 0 {
		fmt.Println("\nOverall Errors:")
		for _, e := range errors {
			fmt.Println(" -", e)
		}
	} else {
		fmt.Println("\nNo overall errors found.")
	}

	branch := make(map[string][][]string)
	for _, row := range rows[1:] {
		if strings.Contains(row[3], "2024A3PS") {
			branch["EEE"] = append(branch["EEE"], row)
		}
		if strings.Contains(row[3], "2024A4PS") {
			branch["MECH"] = append(branch["MECH"], row)
		}
		if strings.Contains(row[3], "2024A5PS") {
			branch["BPHARM"] = append(branch["BPHARM"], row)
		}
		if strings.Contains(row[3], "2024A7PS") {
			branch["CS"] = append(branch["CS"], row)
		}
		if strings.Contains(row[3], "2024A8PS") {
			branch["ENI"] = append(branch["ENI"], row)
		}
		if strings.Contains(row[3], "2024AAPS") {
			branch["ECE"] = append(branch["ECE"], row)
		}
		if strings.Contains(row[3], "2024ADPS") {
			branch["MNC"] = append(branch["MNC"], row)
		}
	}

	for branch, bRows := range branch {
		top3 := Top3(bRows, "quiz")
		fmt.Printf("\nTop 3 students in %s branch for Quiz:\n", branch)
		for i, row := range top3 {
			fmt.Printf("No. %d in %s branch: %s\n", i+1, branch, row[3])
		}
		top3 = Top3(bRows, "midSem")
		fmt.Printf("\nTop 3 students in %s branch for Mid-Sem:\n", branch)
		for i, row := range top3 {
			fmt.Printf("No. %d in %s branch: %s\n", i+1, branch, row[3])
		}
		top3 = Top3(bRows, "labTest")
		fmt.Printf("\nTop 3 students in %s branch for Lab Test:\n", branch)
		for i, row := range top3 {
			fmt.Printf("No. %d in %s branch: %s\n", i+1, branch, row[3])
		}
		top3 = Top3(bRows, "weeklylabs")
		fmt.Printf("\nTop 3 students in %s branch for Weekly Labs:\n", branch)
		for i, row := range top3 {
			fmt.Printf("No. %d in %s branch: %s\n", i+1, branch, row[3])
		}
		top3 = Top3(bRows, "precompre")
		fmt.Printf("\nTop 3 students in %s branch for Pre-Compre:\n", branch)
		for i, row := range top3 {
			fmt.Printf("No. %d in %s branch: %s\n", i+1, branch, row[3])
		}
		top3 = Top3(bRows, "compre")
		fmt.Printf("\nTop 3 students in %s branch for Compre:\n", branch)
		for i, row := range top3 {
			fmt.Printf("No. %d in %s branch: %s\n", i+1, branch, row[3])
		}
		top3 = Top3(bRows, "total")
		fmt.Printf("\nTop 3 students in %s branch for Total:\n", branch)
		for i, row := range top3 {
			fmt.Printf("No. %d in %s branch: %s\n", i+1, branch, row[3])
		}

		c, qAvg, mAvg, ltAvg, lAvg, pAvg, cpAvg, totAvg, errs := computeAverages(bRows)
		fmt.Printf("\nBranch: %s\n", branch)
		fmt.Printf("Records: %d\n", c)
		if c > 0 {
			fmt.Printf("Quiz: %.2f, Mid-Sem: %.2f, Lab Test: %.2f, Weekly Labs: %.2f, Pre-Compre: %.2f, Compre: %.2f, Total: %.2f\n",
				qAvg, mAvg, ltAvg, lAvg, pAvg, cpAvg, totAvg)
		} else {
			fmt.Println("No records for this branch.")
		}
		if len(errs) > 0 {
			fmt.Println("Errors:")
			for _, e := range errs {
				fmt.Println(" -", e)
			}
		} else {
			fmt.Println("No errors found for this branch.")
		}
	}
}
