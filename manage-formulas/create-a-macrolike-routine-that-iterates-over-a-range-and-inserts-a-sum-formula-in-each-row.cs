// Title: Insert Row‑Specific SUM Formulas Across a Range with Aspose.Cells for .NET
// Description: Creates a workbook, fills columns A‑C with sample data, then loops through rows 1‑5 to place a row‑specific =SUM(Ax:Cx) formula in column D using SetFormula and FormulaParseOptions, recalculates the sheet, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | SetFormula | SUM formula | row loop | Excel automation | macro‑like routine | FormulaParseOptions
// Common Searches: add SUM formula to each row Aspose.Cells C# | loop insert formulas Excel using Aspose.Cells | programmatic row totals with Aspose.Cells | set formula per row Aspose.Cells .NET
// Developer Intent: Programmatically add a SUM formula to every row in a target column.
// Use Cases: Generate a totals column for a data table by summing adjacent cells per row. | Build dynamic reports where each row’s subtotal is calculated during workbook creation. | Automate financial worksheets that require row‑level aggregation without manual entry.
// AI Prompts: Write a C# method that receives a worksheet, start/end rows, source column range, and target column, then inserts a row‑specific SUM formula for each row using Aspose.Cells. | Show how to apply locale‑aware SUM formulas inside a loop with FormulaParseOptions in Aspose.Cells. | Explain how to extend the routine to sum a variable number of columns based on a method parameter.

using System;
using Aspose.Cells;

// Creates a workbook, fills columns A‑C with sample data, then loops through rows 1‑5 to place a row‑specific =SUM(Ax:Cx) formula in column D using SetFormula and FormulaParseOptions, recalculates the sheet, and saves the file.
class MacroLikeSumRoutine
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data in columns A, B, and C (rows 1‑5)
        for (int i = 0; i < 5; i++)
        {
            cells[i, 0].PutValue(i + 1);          // Column A
            cells[i, 1].PutValue((i + 1) * 10);   // Column B
            cells[i, 2].PutValue((i + 1) * 100);  // Column C
        }

        // Define the target column (D) where the SUM formulas will be placed
        int targetColumnIndex = 3; // D
        int firstDataRow = 0;      // zero‑based index for row 1
        int lastDataRow = 4;       // zero‑based index for row 5

        // Use default formula parse options
        FormulaParseOptions parseOptions = new FormulaParseOptions();

        // Iterate over each row in the range and set a row‑specific SUM formula
        for (int row = firstDataRow; row <= lastDataRow; row++)
        {
            // Build a formula that sums columns A‑C of the current row
            string formula = $"=SUM(A{row + 1}:C{row + 1})";

            // Apply the formula to the cell in column D of the current row
            cells[row, targetColumnIndex].SetFormula(formula, parseOptions);
        }

        // Recalculate all formulas so that the results are stored
        workbook.CalculateFormula();

        // Save the workbook
        workbook.Save("MacroSumDemo.xlsx");
    }
}
