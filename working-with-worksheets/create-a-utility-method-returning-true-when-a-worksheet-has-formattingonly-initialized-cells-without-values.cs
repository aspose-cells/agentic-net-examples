// Title: C# Method to Detect Formatting‑Only Cells in an Aspose.Cells Worksheet
// Description: Provides WorksheetHelper.HasFormattingOnlyCells, which scans the used range (MaxDataRow/MaxDataColumn) and returns true when a cell has a style applied (IsStyleSet) but contains no value or formula. Handles null cells and empty sheets efficiently.
// Keywords: Aspose.Cells formatting only cells | C# detect styled cells without data | worksheet empty styled cells .NET | check cell style no value Aspose | Excel formatting‑only detection C#
// Common Searches: detect formatting only cells Aspose.Cells C# | find styled cells with no data in .NET Excel | worksheet has cells with style but no value | C# method to check for empty formatted cells | Aspose.Cells skip cells that only have formatting
// Developer Intent: Identify whether any cell in a worksheet has formatting applied but no value or formula.
// Use Cases: Skip saving or processing worksheets that contain only formatted cells to reduce file size. | Validate Excel templates to ensure stray styled cells are removed before generating reports. | Clean up formatting‑only cells prior to converting workbooks to PDF or other output formats.
// AI Prompts: Create a comprehensive unit‑test suite for WorksheetHelper.HasFormattingOnlyCells covering empty sheets, formatting‑only cells, and cells with both style and data. | Refactor HasFormattingOnlyCells to use a Cells iterator with early exit while preserving performance. | Generate sample code that logs the addresses of all formatting‑only cells found in a worksheet using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsUtilities
{
    // Provides WorksheetHelper.HasFormattingOnlyCells, which scans the used range (MaxDataRow/MaxDataColumn) and returns true when a cell has a style applied (IsStyleSet) but contains no value or formula. Handles null cells and empty sheets efficiently.
    public static class WorksheetHelper
    {
        public static bool HasFormattingOnlyCells(Worksheet worksheet)
        {
            Cells cells = worksheet.Cells;

            // Determine the used range to avoid iterating over the entire sheet.
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // If the sheet is completely empty, there are no formatting‑only cells.
            if (maxRow < 0 || maxCol < 0)
                return false;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Skip cells that were never created.
                    if (cell == null)
                        continue;

                    bool styleSet = cell.IsStyleSet;
                    bool hasValueOrFormula = !string.IsNullOrEmpty(cell.StringValue) || cell.IsFormula;

                    // Formatting‑only cell: style is set but no value/formula.
                    if (styleSet && !hasValueOrFormula)
                        return true;
                }
            }

            return false;
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Apply style to a cell without putting any value.
            Style style = workbook.CreateStyle();
            style.Font.Color = System.Drawing.Color.Red;
            cells["B2"].SetStyle(style); // Formatting‑only cell

            // Apply style and value to another cell (should not be counted).
            cells["C3"].SetStyle(style);
            cells["C3"].PutValue("Data");

            bool hasFormattingOnly = WorksheetHelper.HasFormattingOnlyCells(sheet);
            Console.WriteLine("Worksheet has formatting‑only cells: " + hasFormattingOnly);

            // Save the workbook (lifecycle save)
            workbook.Save("FormattingOnlyDemo.xlsx");
        }
    }
}
