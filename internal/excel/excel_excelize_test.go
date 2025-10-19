package excel

import (
	"os"
	"testing"

	"github.com/xuri/excelize/v2"
)

// TestArrayFormulaDetection tests the requiresArrayFormulaType function
func TestArrayFormulaDetection(t *testing.T) {
	tests := []struct {
		name     string
		formula  string
		expected bool
	}{
		{
			name:     "XLOOKUP with array multiplication",
			formula:  "=XLOOKUP(1,(DataTable[MinValue]<=B10)*(DataTable[MaxValue]>=B10),DataTable[Result],\"Not Found\")",
			expected: true,
		},
		{
			name:     "SUMPRODUCT with array operations and ROW",
			formula:  "=SUMPRODUCT(((Data!$A$2:$A$20=$B$1)*(Data!$B$2:$B$20=$B$2))*ROW(Data!$A$2:$A$20))",
			expected: true,
		},
		{
			name:     "INDEX/MATCH with array multiplication",
			formula:  "=INDEX(Data!$C$2:$C$20,MATCH(1,(Data!$A$2:$A$20=$B$1)*(Data!$B$2:$B$20=$B$2),0))",
			expected: true,
		},
		{
			name:     "MATCH with concatenated arrays",
			formula:  "=MATCH($B$3&\"-\"&$B$4,Data!$A:$A&\"-\"&Data!$B:$B,0)",
			expected: true,
		},
		{
			name:     "XLOOKUP with multiple array criteria",
			formula:  "=XLOOKUP(1, (Table[MinValue]<=B10)*(Table[MaxValue]>=B10)*(Table[StartDate]<=B11)*(Table[EndDate]>=B11), Table[Result], \"Not Found\")",
			expected: true,
		},
		{
			name:     "Simple INDEX/MATCH should NOT be array",
			formula:  "=INDEX(Data!C:C, MATCH(B1, Data!A:A, 0))",
			expected: false,
		},
		{
			name:     "Simple SUM should NOT be array",
			formula:  "=SUM(A1:A10)",
			expected: false,
		},
		{
			name:     "Complex formula without array ops should NOT be array",
			formula:  "=IF(B5, SQRT((B12/100*A28)^2+(B13/100*B14*1000)^2), 0)",
			expected: false,
		},
		{
			name:     "Simple multiplication should NOT be array",
			formula:  "=A1*B1",
			expected: false,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := requiresArrayFormulaType(tt.formula)
			if result != tt.expected {
				t.Errorf("requiresArrayFormulaType(%q) = %v, want %v", tt.formula, result, tt.expected)
			}
		})
	}
}

// TestArrayFormulaIntegration tests that array formulas are written and read correctly
func TestArrayFormulaIntegration(t *testing.T) {
	testFile := "/tmp/test_array_formula_integration.xlsm"
	defer os.Remove(testFile)

	// Create a new Excel file
	f := excelize.NewFile()
	defer f.Close()

	// Create data sheet
	dataSheet := "Data"
	dataIdx, _ := f.NewSheet(dataSheet)
	f.SetActiveSheet(dataIdx)

	// Add test data
	f.SetCellValue(dataSheet, "A1", "MinValue")
	f.SetCellValue(dataSheet, "B1", "MaxValue")
	f.SetCellValue(dataSheet, "C1", "Result")
	f.SetCellValue(dataSheet, "A2", 0)
	f.SetCellValue(dataSheet, "B2", 10)
	f.SetCellValue(dataSheet, "C2", "Low")
	f.SetCellValue(dataSheet, "A3", 11)
	f.SetCellValue(dataSheet, "B3", 20)
	f.SetCellValue(dataSheet, "C3", "Medium")

	f.DeleteSheet("Sheet1")

	// Save initial file
	if err := f.SaveAs(testFile); err != nil {
		t.Fatalf("Failed to create test file: %v", err)
	}
	f.Close()

	// Open via our Excel interface
	workbook, closeFn, err := OpenFile(testFile)
	if err != nil {
		t.Fatalf("Failed to open file: %v", err)
	}
	defer closeFn()

	// Create test sheet
	if err := workbook.CreateNewSheet("Test"); err != nil {
		t.Fatalf("Failed to create test sheet: %v", err)
	}

	worksheet, err := workbook.FindSheet("Test")
	if err != nil {
		t.Fatalf("Failed to find test sheet: %v", err)
	}
	defer worksheet.Release()

	// Add test formulas
	testCases := []struct {
		cell    string
		formula string
	}{
		{
			cell:    "B1",
			formula: "=XLOOKUP(1,(Data!A$2:A$3<=15)*(Data!B$2:B$3>=15),Data!C$2:C$3,\"Not Found\")",
		},
		{
			cell:    "B2",
			formula: "=INDEX(Data!C$2:C$3,MATCH(1,(Data!A$2:A$3<=15)*(Data!B$2:B$3>=15),0))",
		},
		{
			cell:    "B3",
			formula: "=SUM(A1:A10)",
		},
	}

	for _, tc := range testCases {
		if err := worksheet.SetFormula(tc.cell, tc.formula); err != nil {
			t.Errorf("Failed to set formula in %s: %v", tc.cell, err)
		}
	}

	// Save
	if err := workbook.Save(); err != nil {
		t.Fatalf("Failed to save file: %v", err)
	}

	// Close and reopen
	closeFn()

	workbook2, closeFn2, err := OpenFile(testFile)
	if err != nil {
		t.Fatalf("Failed to reopen file: %v", err)
	}
	defer closeFn2()

	worksheet2, err := workbook2.FindSheet("Test")
	if err != nil {
		t.Fatalf("Failed to find test sheet on reopen: %v", err)
	}
	defer worksheet2.Release()

	// Verify formulas are preserved
	for _, tc := range testCases {
		formula, err := worksheet2.GetFormula(tc.cell)
		if err != nil {
			t.Errorf("Failed to read formula from %s: %v", tc.cell, err)
			continue
		}

		if formula != tc.formula {
			t.Errorf("Formula mismatch in %s:\nExpected: %s\nGot:      %s", tc.cell, tc.formula, formula)
		}
	}
}
