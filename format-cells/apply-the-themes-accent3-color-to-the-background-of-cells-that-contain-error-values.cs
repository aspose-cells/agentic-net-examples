// Title: Apply Theme Accent3 Background to Error Cells in C# with Aspose.Cells
// Description: This C# example demonstrates how to scan a worksheet, detect cells that contain error values (e.g., #DIV/0!, #NAME?), and apply the workbook's Accent3 theme color as a solid background using Aspose.Cells for .NET. The code creates a style with ThemeColorType.Accent3, assigns it to each error cell, and saves the result.
// Keywords: Aspose.Cells C# error highlighting | theme accent color background | Accent3 cell style | detect cell errors Aspose | .NET spreadsheet formatting | CellRichValue error detection | apply theme color to cells | Excel error cell background | Aspose.Cells sample code
// Common Searches: how to color error cells with Aspose.Cells | apply theme Accent3 background to Excel cells C# | detect #DIV/0! and set background in Aspose.Cells | C# code to highlight cells with errors using theme colors | Aspose.Cells example for error‑value formatting
// Developer Intent: Automatically highlight any cell that returns an error by applying the workbook’s Accent3 theme color as the cell background.
// Use Cases: Mark calculation errors in financial models so reviewers can spot them instantly. | Flag invalid formulas in data‑validation sheets before publishing to end users. | Create error‑aware dashboards where error cells are shaded with a consistent theme color.
// AI Prompts: Generate C# code that uses Aspose.Cells to set the Accent3 theme color as the background for all error cells in a worksheet. | Show an alternative approach that uses conditional formatting instead of a manual loop to highlight error values with Accent3. | Explain performance‑optimised techniques for scanning large worksheets while applying theme‑based styling to error cells.

using System;
using System.Drawing;
using Aspose.Cells;

namespace ApplyAccent3ToErrorCells
{
    // This C# example demonstrates how to scan a worksheet, detect cells that contain error values (e.g., #DIV/0!, #NAME?), and apply the workbook's Accent3 theme color as a solid background using Aspose.Cells for .NET. The code creates a style with ThemeColorType.Accent3, assigns it to each error cell, and saves the result.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Example data: create some error cells
            cells["A1"].Formula = "=1/0";          // #DIV/0! error
            cells["B2"].Formula = "=UNKNOWNFUNC()"; // #NAME? error
            cells["C3"].PutValue("Normal text");   // non‑error cell

            // Determine the used range of the worksheet
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Iterate through each cell in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Retrieve the rich value to check for an error
                    CellRichValue richValue = cell.GetRichValue();

                    // If the cell contains an error (ErrorValue enum is non‑zero)
                    if (richValue != null && richValue.ErrorValue != 0)
                    {
                        // Create a style that uses the theme's Accent3 color as background
                        Style errorStyle = workbook.CreateStyle();
                        errorStyle.Pattern = BackgroundType.Solid;
                        errorStyle.BackgroundThemeColor = new ThemeColor(ThemeColorType.Accent3, 0);

                        // Apply the style to the error cell
                        cell.SetStyle(errorStyle);
                    }
                }
            }

            // Save the workbook
            workbook.Save("ErrorCellsWithAccent3Background.xlsx");
        }
    }
}
