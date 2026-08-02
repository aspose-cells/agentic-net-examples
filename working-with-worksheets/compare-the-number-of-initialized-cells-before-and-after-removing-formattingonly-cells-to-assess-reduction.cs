// Title: C# Aspose.Cells Example – Compare instantiated cell count before and after removing formatting‑only cells
// Description: Creates a workbook, adds data and style‑only cells, records CountLarge, resets empty cells with custom styles to the default, removes unused styles, then shows the before/after cell count to illustrate memory savings.
// Keywords: Aspose.Cells CountLarge | formatting‑only cells | remove unused styles | C# spreadsheet optimization | cell instantiation reduction | Aspose.Cells memory cleanup | worksheet cell count comparison
// Common Searches: Aspose.Cells count cells before after removing formatting only | how to delete empty styled cells in Aspose.Cells C# | reduce workbook size by cleaning up styles Aspose.Cells | C# example for resetting cell style to default Aspose.Cells | compare instantiated cells Aspose.Cells
// Developer Intent: Show how resetting the style of empty cells removes formatting‑only cells and lowers the instantiated cell count reported by CountLarge.
// Use Cases: Quantify memory and performance impact of cleaning up style‑only cells in large Excel files. | Validate that ResetStyle on empty cells decreases the CountLarge value of a worksheet. | Generate a quick before/after report for workbook optimization audits.
// AI Prompts: Generate a C# method that scans a worksheet, clears formatting‑only cells, and returns the CountLarge values before and after the cleanup using Aspose.Cells. | Explain why resetting a cell's style to the default removes a formatting‑only cell and how this affects the CountLarge property. | Suggest scalable techniques for detecting and removing formatting‑only cells in worksheets with millions of rows in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds data and style‑only cells, records CountLarge, resets empty cells with custom styles to the default, removes unused styles, then shows the before/after cell count to illustrate memory savings.
    public class FormattingOnlyCellsComparison
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // -------------------------------------------------
            // 1. Add regular data cells
            // -------------------------------------------------
            cells["A1"].PutValue("Header");
            cells["A2"].PutValue("Data1");
            cells["A3"].PutValue("Data2");

            // -------------------------------------------------
            // 2. Add formatting‑only cells (no value, only style)
            // -------------------------------------------------
            Style fmtStyle = wb.CreateStyle();
            fmtStyle.Font.IsBold = true;
            fmtStyle.Font.Color = Color.Red;

            cells["B1"].SetStyle(fmtStyle); // formatting only
            cells["C2"].SetStyle(fmtStyle); // formatting only
            cells["D4"].SetStyle(fmtStyle); // formatting only (outside data range)

            // -------------------------------------------------
            // 3. Count instantiated cells before cleanup
            // -------------------------------------------------
            long countBefore = cells.CountLarge;
            Console.WriteLine($"Instantiated cells before removing formatting‑only cells: {countBefore}");

            // -------------------------------------------------
            // 4. Remove formatting‑only cells
            //    If a cell has no value but has a custom style, reset its style to the default.
            // -------------------------------------------------
            // Define a reasonable scan area that includes possible formatting‑only cells.
            int maxRow = Math.Max(cells.MaxDataRow, 10);
            int maxCol = Math.Max(cells.MaxDataColumn, 10);

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];
                    if (cell == null) continue; // cell not instantiated yet

                    bool hasValue = cell.Value != null && !(cell.Value is string s && string.IsNullOrEmpty(s));
                    bool hasCustomStyle = !cell.GetStyle().Equals(wb.DefaultStyle);

                    if (!hasValue && hasCustomStyle)
                    {
                        // Reset to default style, effectively removing the formatting‑only cell
                        cell.SetStyle(wb.DefaultStyle);
                    }
                }
            }

            // Remove any styles that are no longer used after the cleanup
            wb.RemoveUnusedStyles();

            // -------------------------------------------------
            // 5. Count instantiated cells after cleanup
            // -------------------------------------------------
            long countAfter = cells.CountLarge;
            Console.WriteLine($"Instantiated cells after removing formatting‑only cells: {countAfter}");

            // -------------------------------------------------
            // 6. Report reduction
            // -------------------------------------------------
            Console.WriteLine($"Reduction in instantiated cells: {countBefore - countAfter}");

            // -------------------------------------------------
            // 7. Save the workbook for visual verification
            // -------------------------------------------------
            wb.Save("FormattingOnlyCellsComparison.xlsx");
        }
    }
}
