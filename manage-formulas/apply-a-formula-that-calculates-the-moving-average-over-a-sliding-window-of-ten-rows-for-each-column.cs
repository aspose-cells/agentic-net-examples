// Title: Insert a 10‑row sliding‑window moving average formula into every column of an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that writes an AVERAGE formula using OFFSET to compute a 10‑row moving average for each column in a worksheet. | Show how to trigger a full workbook recalculation after setting the moving‑average formulas and then save the file as an .xlsx document. | Adapt the example so the window size is passed as a parameter and the formula is applied only to columns A through D.
// Common Searches: aspnet c# how to programmatically add a 10‑row moving average with Aspose.Cells | using OFFSET in Aspose.Cells to create dynamic average range for each column | recalculate all formulas in an Aspose.Cells workbook after inserting formulas | save workbook as xlsx after applying sliding window calculations with Aspose.Cells
// Tags: apply moving average formula Aspose.Cells C# | OFFSET based dynamic range Excel Aspose.Cells | recalculate workbook formulas Aspose.Cells | save workbook as xlsx Aspose.Cells | set formula for multiple columns programmatically

using System;
using Aspose.Cells;

// The example creates a workbook, fills columns A‑D with sample data, inserts an AVERAGE formula built with OFFSET that calculates a 10‑row moving average for each column, forces a full recalculation, and saves the result as MovingAverageResult.xlsx.
class MovingAverageExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data in columns A to D (for demonstration)
        // Assume data starts at row 1 (index 0) and goes down to row 30
        for (int row = 0; row < 30; row++)
        {
            cells[row, 0].PutValue(row + 1);               // Column A
            cells[row, 1].PutValue((row + 1) * 2);         // Column B
            cells[row, 2].PutValue((row + 1) * 3);         // Column C
            cells[row, 3].PutValue((row + 1) * 4);         // Column D
        }

        // Define the window size for the moving average
        int windowSize = 10;

        // Apply moving average formula for each column
        // The result will be placed in the same column, starting from the row where the first full window is available
        // (i.e., row index = windowSize - 1)
        for (int col = 0; col < 4; col++) // Adjust column count as needed
        {
            for (int row = windowSize - 1; row < 30; row++)
            {
                // Build the AVERAGE formula using OFFSET to create a dynamic range of 'windowSize' rows ending at the current row
                // Formula example for column A, row 11 (1‑based): =AVERAGE(OFFSET(A1,ROW()-10,0,10,1))
                string columnLetter = CellsHelper.ColumnIndexToName(col);
                string startCell = $"{columnLetter}1";
                string formula = $"AVERAGE(OFFSET({startCell},ROW()-{windowSize},0,{windowSize},1))";

                // Set the formula in the target cell (row and column are zero‑based)
                cells[row, col].Formula = formula;
            }
        }

        // Recalculate all formulas so that the workbook contains the computed values
        workbook.CalculateFormula();

        // Save the workbook to a file
        workbook.Save("MovingAverageResult.xlsx");
    }
}
