package main

import (
	"fmt"
	"math"
	"strconv"

	//"encoding/json"
	"flag"

	"github.com/xuri/excelize/v2"
)

func main() {
	pathFlag := flag.String("path", "", "Path to the file")
	jsonFlag := flag.String("export", "", "Export to JSON")
	classFlag := flag.String("class", "", "Class to filter")
	flag.Parse()
	fmt.Println(*jsonFlag)
	if *pathFlag == "" {
		fmt.Println("Please provide the file path")
		return
	}
	f, err := excelize.OpenFile(*pathFlag)
	if err != nil {
		fmt.Println(err)
		return
	}
	if *classFlag != "" {
		fmt.Println("Filtering for class", *classFlag)
	}
	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	// Get all the rows in the Sheet1.
	rows, err := f.GetRows(f.WorkBook.Sheets.Sheet[0].Name)
	if err != nil {
		fmt.Println(err)
		return
	}
	total := 0.00
	count := 0

	// Component wise map element
	exams := make(map[int]float64)
	components := make(map[int][2][]interface{})
	branches := make(map[string]float64)
	branch_count := make(map[string]int)
	branch := ""
	emplid := ""
	headers := rows[0]
	//fmt.Println("Headers:", headers)

	for _, row := range rows[1:] {
		if len(row) == 0 {
			continue
		}
		sum := 0.00
		emplid = row[2]
		if *classFlag != "" && row[1] != *classFlag {
			continue
		}
		for rowNo, colCell := range row {
			//fmt.Print(colCell, "\t")
			if rowNo == 3 {
				if colCell[:4] != "2024" {
					branch = ""
				} else {
					branch = colCell[4:6]
					branch_count[branch]++
				}
			}
			// Rows with numbers
			if rowNo >= 4 && rowNo <= 9 {
				i, err := strconv.ParseFloat(colCell, 32)
				if err != nil {
					continue
				}
				if rowNo != 8 {
					sum += i
				}
				// Component wise sum
				exams[rowNo] += i
				components[rowNo] = [2][]interface{}{append(components[rowNo][0], i), append(components[rowNo][1], emplid)}
			}
			// Total
			if rowNo == 10 {
				i, _ := strconv.ParseFloat(colCell, 32)
				i = math.Round(i*100) / 100
				sum = math.Round(sum*100) / 100
				if sum != i {
					fmt.Println("Discrepancy Found! Total not matching")
					fmt.Println("For", emplid, "Expected total to be", i, "but turned out to be", sum)
				}
				total += sum
				count++
				if branch != "" {
					branches[branch] += sum
				}
				sum = 0
			}
		}
		//fmt.Println()
	}
	fmt.Println("Averages:")
	fmt.Println("Total: ", total/float64(count))

	// Component wise avs
	for i, v := range exams {
		fmt.Println(headers[i], ":", v/float64(count))
	}

	// Branch wise avs
	fmt.Println("24 Batch Single Degree Branch wise averages:")
	for i, v := range branches {
		fmt.Println(i, ":", v/float64(branch_count[i]))
	}
	for rowNo, header := range headers[4:10] {
		rowNo += 4
		top3, top3emplids := gettop3(components[rowNo][0], components[rowNo][1])
		fmt.Println("Top 3 in", header, ":")
		for i, v := range top3 {
			fmt.Println(i+1, top3emplids[i], ":", v)
		}
	}
}

func gettop3(marks []interface{}, emplids []interface{}) ([]float64, []string) {
	// Get the top 3 marks
	top3 := []float64{0, 0, 0}
	top3emplids := []string{"", "", ""}
	for i, mark := range marks {
		mark := mark.(float64)
		empli := emplids[i].(string)

		if mark > top3[0] {
			top3[2] = top3[1]
			top3[1] = top3[0]
			top3[0] = mark
			top3emplids[2] = top3emplids[1]
			top3emplids[1] = top3emplids[0]
			top3emplids[0] = empli
		} else if mark > top3[1] {
			top3[2] = top3[1]
			top3[1] = mark
			top3emplids[2] = top3emplids[1]
			top3emplids[1] = empli
		} else if mark > top3[2] {
			top3[2] = mark
			top3emplids[2] = empli
		}
	}
	return top3, top3emplids
}
