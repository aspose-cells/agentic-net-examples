using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPrintAreaDemo
{
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Create a new workbook and populate some sample data.
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Fill cells A1:C5 with sample values.
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // ------------------------------------------------------------
            // 2. Define a print area (e.g., B2:C4) and save the workbook
            //    exporting only that area to HTML.
            // ------------------------------------------------------------
            sheet.PageSetup.PrintArea = "B2:C4";

            HtmlSaveOptions optionsPrintAreaOnly = new HtmlSaveOptions
            {
                ExportPrintAreaOnly = true, // Export only the defined print area.
                ExportGridLines = true
            };
            workbook.Save("PrintAreaOnly.html", optionsPrintAreaOnly);

            Console.WriteLine("Saved PrintAreaOnly.html – contains only B2:C4.");

            // ------------------------------------------------------------
            // 3. Remove the defined print area.
            //    Setting PrintArea to an empty string clears the setting.
            // ------------------------------------------------------------
            sheet.PageSetup.PrintArea = string.Empty; // No explicit print area now.

            // ------------------------------------------------------------
            // 4. After removing the print area, the default printing range
            //    becomes the *used range* of the worksheet (the smallest
            //    rectangle that contains all non‑empty cells). In this
            //    example the used range is A1:C5.
            // ------------------------------------------------------------
            // Verify the used range via Cells.MaxDisplayRange.
            var usedRange = sheet.Cells.MaxDisplayRange; // Returns a Range object
            int startRow = usedRange.FirstRow;
            int startCol = usedRange.FirstColumn;
            int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
            int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;
            Console.WriteLine($"Used range after clearing PrintArea: {startRow},{startCol} to {endRow},{endCol}");

            // ------------------------------------------------------------
            // 5. Save the workbook again without ExportPrintAreaOnly.
            //    The whole used range (A1:C5) will be exported.
            // ------------------------------------------------------------
            HtmlSaveOptions optionsFull = new HtmlSaveOptions
            {
                ExportPrintAreaOnly = false, // Export the entire used range.
                ExportGridLines = true
            };
            workbook.Save("FullSheet.html", optionsFull);

            Console.WriteLine("Saved FullSheet.html – contains the full used range (A1:C5).");
        }
    }
}