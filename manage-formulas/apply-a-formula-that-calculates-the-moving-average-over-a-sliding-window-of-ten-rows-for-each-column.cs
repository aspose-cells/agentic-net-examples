// Title: Add a 10‑row moving average formula to every column in an Excel worksheet with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells example to set a moving average formula for each column | how to create a sliding‑window average over rows using Aspose.Cells workbook | apply dynamic AVERAGE range for the last 10 rows in Excel with Aspose.Cells | calculate rolling average in an Excel file programmatically with Aspose.Cells .NET | populate worksheet and assign moving average formulas using Aspose.Cells C#
// Tags: Aspose.Cells moving average formula C# | sliding window AVERAGE range Aspose.Cells | dynamic Excel range formula Aspose.Cells | calculate workbook formulas Aspose.Cells | save Excel file with moving average Aspose.Cells

using System;
using Aspose.Cells;

namespace MovingAverageExample
{
    // The C# program creates a new workbook, fills columns A‑D with values 1‑100, determines the used range, and then iterates over each column. Starting from the 10th row, it assigns an AVERAGE formula that references the preceding ten rows, producing a rolling average for every column. After calculating all formulas, the workbook is saved as MovingAverageResult.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // OPTIONAL: Populate the worksheet with sample data for demonstration.
            // Here we fill columns A to D with values 1..100.
            for (int col = 0; col < 4; col++)
            {
                for (int row = 0; row < 100; row++)
                {
                    sheet.Cells[row, col].PutValue(row + 1);
                }
            }

            // Determine the used range of the worksheet.
            int maxRow = sheet.Cells.MaxDataRow;       // zero‑based index of the last row with data
            int maxColumn = sheet.Cells.MaxDataColumn; // zero‑based index of the last column with data

            // Apply moving average formula with a sliding window of 10 rows.
            // The formula is placed starting from the 10th row (index 9) because
            // the first nine rows do not have enough preceding data.
            for (int col = 0; col <= maxColumn; col++)
            {
                // Convert column index to Excel column letter (e.g., 0 -> "A").
                string columnLetter = CellsHelper.ColumnIndexToName(col);

                for (int row = 9; row <= maxRow; row++)
                {
                    // Excel rows are 1‑based, so add 1 to the zero‑based index.
                    int currentExcelRow = row + 1;

                    // Start of the 10‑row window.
                    int startExcelRow = currentExcelRow - 9;

                    // Build the range string, e.g., "A1:A10".
                    string range = $"{columnLetter}{startExcelRow}:{columnLetter}{currentExcelRow}";

                    // Set the AVERAGE formula for the moving average.
                    sheet.Cells[row, col].Formula = $"AVERAGE({range})";
                }
            }

            // Calculate all formulas in the workbook.
            workbook.CalculateFormula();

            // Save the workbook to a file.
            workbook.Save("MovingAverageResult.xlsx");
        }
    }
}
