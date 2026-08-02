// Title: C# – Hide Zero Values in Rows > 100 Using Aspose.Cells
// Description: Creates a workbook, populates 120 rows, defines a custom number format "0;-0;;@" that suppresses zero display, and applies this style only to numeric zeros in rows after the 100th row, then saves the file as XLSX.
// Keywords: Aspose.Cells C# hide zeros | custom number format zero suppression | apply style after row 100 | Excel zero display control | C# workbook formatting Aspose
// Common Searches: Aspose.Cells hide zero values after row 100 | C# custom number format to hide zeros in Excel | apply style to specific rows Aspose.Cells | suppress zero display in large worksheets C# | conditional formatting zeros beyond row 100 Aspose
// Developer Intent: Apply a style that hides numeric zeros only in rows beyond the 100th row while keeping earlier rows unchanged.
// Use Cases: Financial statements where summary rows show zeros but detailed rows should appear blank. | Large data exports where trailing rows contain many zero values that clutter the view. | Templates that automatically conceal zeros in the lower section of a sheet for cleaner presentation.
// AI Prompts: Generate Aspose.Cells C# code to hide zero values in rows greater than a given index using a custom number format. | Show how to use conditional formatting in Aspose.Cells to suppress zeros after row 100. | Explain how to adapt the style‑application loop for multiple worksheets and a configurable start row.

using System;
using Aspose.Cells;

namespace HideZeroValuesBeyondRow100
{
    // Creates a workbook, populates 120 rows, defines a custom number format "0;-0;;@" that suppresses zero display, and applies this style only to numeric zeros in rows after the 100th row, then saves the file as XLSX.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // Sample data creation (optional, for demonstration purposes)
            // ------------------------------------------------------------
            // Fill first 120 rows with some numeric values, including zeros
            for (int row = 0; row < 120; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    // Insert zero in every third column for rows beyond 100
                    double value = (row >= 100 && col == 2) ? 0 : row + col + 1;
                    cells[row, col].PutValue(value);
                }
            }

            // ------------------------------------------------------------
            // Create a style that hides zero values using a custom format
            // Format sections: positive;negative;zero;text
            // The third section is left empty to hide zeros
            // ------------------------------------------------------------
            Style hideZeroStyle = workbook.CreateStyle();
            hideZeroStyle.Custom = "0;-0;;@";

            // ------------------------------------------------------------
            // Apply the style only to zero-valued cells in rows > 100
            // ------------------------------------------------------------
            int startRow = 100; // zero‑based index (row 101 in Excel)
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int r = startRow; r <= maxRow; r++)
            {
                for (int c = 0; c <= maxCol; c++)
                {
                    Cell cell = cells[r, c];
                    // Check if the cell contains a numeric zero
                    if (cell.Type == CellValueType.IsNumeric && cell.DoubleValue == 0)
                    {
                        cell.SetStyle(hideZeroStyle);
                    }
                }
            }

            // Save the workbook
            workbook.Save("HideZeroValuesBeyondRow100.xlsx", SaveFormat.Xlsx);
        }
    }
}
