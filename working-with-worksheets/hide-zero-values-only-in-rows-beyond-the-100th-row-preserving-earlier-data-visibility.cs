// Title: C# AspNet Aspose.Cells Example: Hide Zero Values Only After Row 100
// Description: Shows how to programmatically suppress numeric zeros in rows 101 and beyond with Aspose.Cells for .NET. The sample creates a workbook, populates rows 1‑150, applies the custom format "0;-0;;@" to each zero cell after the 100th row, and saves the result as an XLSX file.
// Keywords: Aspose.Cells | C# | .NET | hide zero values | custom number format | zero suppression | Excel row 101 | conditional formatting Aspose | programmatic Excel styling | GitHub Aspose.Cells example | worksheet style zero display | large worksheet performance
// Common Searches: Aspose.Cells hide zeros after row 100 C# | custom number format to hide zero values in Excel using .NET | programmatically hide zero cells beyond a specific row Aspose | C# example for conditional zero display in large worksheets | Aspose.Cells hide zeros in rows 101+
// Developer Intent: Apply a style that hides numeric zero cells only in rows beyond the 100th row while leaving earlier zeros visible.
// Use Cases: Financial statements where summary rows (first 100) must show zeros but detailed rows should appear cleaner. | Exporting massive datasets and removing trailing zero clutter after a defined row threshold. | Creating region‑specific worksheets that keep global DisplayZeros enabled but suppress zeros in a designated range.
// AI Prompts: Generate C# code with Aspose.Cells that hides zero values only from row 101 onward using a custom number format. | Explain how to iterate from the 101st row to the last data row and apply a style that suppresses zeros without affecting earlier rows. | Show how to modify the custom format string to hide zeros while preserving positive and negative numbers in Aspose.Cells.

using System;
using Aspose.Cells;

namespace HideZeroValuesBeyondRow100
{
    // Shows how to programmatically suppress numeric zeros in rows 101 and beyond with Aspose.Cells for .NET. The sample creates a workbook, populates rows 1‑150, applies the custom format "0;-0;;@" to each zero cell after the 100th row, and saves the result as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for demonstration (rows 1 to 150, columns A to C)
            for (int row = 0; row < 150; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    // Insert zero values at every 5th row for testing
                    if ((row + 1) % 5 == 0)
                        cells[row, col].PutValue(0);
                    else
                        cells[row, col].PutValue(row + col + 1);
                }
            }

            // Ensure zeros are displayed globally (default behavior)
            sheet.DisplayZeros = true;

            // Define the starting row index (zero‑based) for rows beyond the 100th row
            int startRowIndex = 100; // corresponds to Excel row 101

            // Determine the last row that contains data
            int lastDataRow = cells.MaxDataRow;

            // Loop through rows beyond the 100th row
            for (int i = startRowIndex; i <= lastDataRow; i++)
            {
                // Loop through all columns that contain data in the current row
                int lastDataColumn = cells.MaxDataColumn;
                for (int j = 0; j <= lastDataColumn; j++)
                {
                    Cell cell = cells[i, j];

                    // Check if the cell holds a numeric zero
                    if (cell.Type == CellValueType.IsNumeric && cell.DoubleValue == 0)
                    {
                        // Apply a custom number format that hides zero values
                        Style hideZeroStyle = workbook.CreateStyle();
                        hideZeroStyle.Custom = "0;-0;;@"; // third section (zero) is empty
                        cell.SetStyle(hideZeroStyle);
                    }
                }
            }

            // Save the workbook to a file
            workbook.Save("HideZerosBeyondRow100.xlsx", SaveFormat.Xlsx);
        }
    }
}
