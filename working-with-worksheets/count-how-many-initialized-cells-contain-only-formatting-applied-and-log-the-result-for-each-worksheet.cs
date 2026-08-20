// Title: Count Formatting‑Only Initialized Cells per Worksheet with Aspose.Cells for .NET
// Description: A C# example that creates a workbook, adds cells with value only, style only, and both, then walks through every worksheet, enumerates only instantiated cells, detects non‑default formatting without a value, logs the count per sheet, and saves the file. Ideal for developers needing to audit or clean up styling in large Excel files.
// Keywords: Aspose.Cells count formatting only cells | initialized cells enumeration .NET | detect non‑default style Aspose.Cells | worksheet formatting audit C# | Excel cell style only detection | Aspose.Cells workbook analysis | C# Excel formatting only cells
// Common Searches: how to count cells with only formatting using Aspose.Cells | Aspose.Cells enumerate instantiated cells for style check | C# find cells that have style but no value in Excel | log formatting‑only cells per worksheet Aspose | detect non‑default cell formatting Aspose.Cells .NET
// Developer Intent: Identify and tally cells that contain styling but no data in each worksheet of an Aspose.Cells workbook.
// Use Cases: Generate a report of styling‑only cells to streamline workbook size before distribution. | Audit multiple sheets for orphaned formats that may affect performance or visual consistency. | Exclude formatting‑only cells from data export pipelines to improve processing speed.
// AI Prompts: Create a function that returns a dictionary of worksheet names and formatting‑only cell counts using Aspose.Cells. | Rewrite the counting loop with LINQ to improve readability and performance. | Add logic to highlight every formatting‑only cell in yellow after the count is logged.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // A C# example that creates a workbook, adds cells with value only, style only, and both, then walks through every worksheet, enumerates only instantiated cells, detects non‑default formatting without a value, logs the count per sheet, and saves the file. Ideal for developers needing to audit or clean up styling in large Excel files.
    class CountFormattingOnlyCells
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Sample data: cell with value only
                cells["A1"].PutValue("Data");

                // Sample data: cell with formatting only (no value)
                Style fmtOnly = workbook.CreateStyle();
                fmtOnly.Font.Color = Color.Red;
                fmtOnly.Font.IsBold = true;
                cells["B2"].SetStyle(fmtOnly);

                // Sample data: cell with both value and formatting (should not be counted)
                Style fmtBoth = workbook.CreateStyle();
                fmtBoth.Font.Color = Color.Blue;
                cells["C3"].PutValue(123);
                cells["C3"].SetStyle(fmtBoth);

                // Iterate through each worksheet in the workbook
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    Cells wsCells = ws.Cells;
                    int formattingOnlyCount = 0;

                    // Enumerate only instantiated cells (initialized cells)
                    foreach (Cell cell in wsCells)
                    {
                        // Determine if the cell has no value
                        bool hasNoValue = cell.Value == null || string.IsNullOrEmpty(cell.StringValue);

                        // Determine if the cell has any non‑default formatting
                        Style cellStyle = cell.GetStyle();
                        bool hasFormatting = HasNonDefaultFormatting(cellStyle);

                        // Count cells that have formatting but no value
                        if (hasNoValue && hasFormatting)
                        {
                            formattingOnlyCount++;
                        }
                    }

                    // Log the result for the current worksheet
                    Console.WriteLine($"Worksheet \"{ws.Name}\": Cells with only formatting = {formattingOnlyCount}");
                }

                // Save the workbook (lifecycle rule: save)
                string outputPath = "FormattedOnlyCellsCount.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Checks whether a style contains any formatting different from the default style
        private static bool HasNonDefaultFormatting(Style style)
        {
            // Font checks
            if (style.Font.Color != Color.Empty) return true;
            if (style.Font.IsBold) return true;
            if (style.Font.IsItalic) return true;
            if (style.Font.Underline != FontUnderlineType.None) return true;
            if (style.Font.Size != 0) return true;

            // Background checks
            if (style.BackgroundColor != Color.Empty) return true;
            if (style.Pattern != BackgroundType.None) return true;

            // Border checks (any border style set)
            if (style.Borders[BorderType.LeftBorder].LineStyle != CellBorderType.None) return true;
            if (style.Borders[BorderType.RightBorder].LineStyle != CellBorderType.None) return true;
            if (style.Borders[BorderType.TopBorder].LineStyle != CellBorderType.None) return true;
            if (style.Borders[BorderType.BottomBorder].LineStyle != CellBorderType.None) return true;

            // If none of the above, consider it default (no formatting)
            return false;
        }
    }
}
