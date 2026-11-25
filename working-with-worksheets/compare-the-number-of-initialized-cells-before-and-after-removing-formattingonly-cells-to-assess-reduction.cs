// Title: Compare initialized cell count before and after clearing formatting‑only cells using Aspose.Cells for .NET
// Common Searches: Aspose.Cells count cells with data before and after removing style‑only cells in C# | C# Aspose.Cells calculate reduction of initialized cells after clearing formatting | How to get number of initialized cells in a worksheet using Aspose.Cells .NET | Remove formatting‑only cells with Aspose.Cells and measure cell count change | Aspose.Cells .NET example to compare cell count pre and post formatting cleanup
// Tags: initialized cell count Aspose.Cells | clear style from empty cells Aspose.Cells | worksheet formatting cleanup Aspose.Cells | cell count reduction percentage .NET | clear formatting from cells without data Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsFormattingRemovalDemo
{
    // The sample loads or creates an Excel workbook, counts cells that contain data or formulas, clears styles from cells that have no value or formula, recounts the initialized cells, computes the absolute and percentage reduction, and saves the cleaned workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Paths to the input and output workbooks.
                string inputPath = "input.xlsx";
                string outputPath = "output_removed_formatting.xlsx";

                // Load an existing workbook or create a sample one.
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                    Console.WriteLine($"Loaded workbook from \"{inputPath}\".");
                }
                else
                {
                    workbook = CreateSampleWorkbook();
                    workbook.Save(inputPath);
                    Console.WriteLine($"Sample workbook created and saved to \"{inputPath}\".");
                }

                // Work with the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Count initialized (data) cells before removing formatting‑only cells.
                int beforeCount = GetInitializedCellCount(sheet);
                Console.WriteLine($"Initialized cells before removing formatting‑only cells: {beforeCount}");

                // Remove formatting from cells that contain no data or formula.
                RemoveFormattingOnlyCells(sheet, workbook);
                Console.WriteLine("Formatting‑only cells have been cleared.");

                // Count initialized cells after the removal.
                int afterCount = GetInitializedCellCount(sheet);
                Console.WriteLine($"Initialized cells after removing formatting‑only cells: {afterCount}");

                // Calculate reduction.
                int reduction = beforeCount - afterCount;
                double reductionPercent = beforeCount == 0 ? 0 : (double)reduction / beforeCount * 100;
                Console.WriteLine($"Reduction: {reduction} cells ({reductionPercent:F2}% decrease)");

                // Save the workbook after formatting removal.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook after formatting removal saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Counts cells that contain a value or a formula.
        static int GetInitializedCellCount(Worksheet sheet)
        {
            int count = 0;
            foreach (Cell cell in sheet.Cells)
            {
                if (cell.Value != null || !string.IsNullOrEmpty(cell.Formula))
                {
                    count++;
                }
            }
            return count;
        }

        // Clears style from cells that have no value and no formula.
        static void RemoveFormattingOnlyCells(Worksheet sheet, Workbook workbook)
        {
            // Default (empty) style.
            Style defaultStyle = workbook.CreateStyle();

            foreach (Cell cell in sheet.Cells)
            {
                if (cell.Value == null && string.IsNullOrEmpty(cell.Formula))
                {
                    cell.SetStyle(defaultStyle);
                }
            }
        }

        static Workbook CreateSampleWorkbook()
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // Data cells
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Price");
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(1.2);
            cells["A3"].PutValue("Banana");
            cells["B3"].PutValue(0.8);

            // Formula cell
            cells["A4"].PutValue("Total");
            cells["B4"].Formula = "=SUM(B2:B3)";

            // Formatting‑only cells (no data, just style)
            Style fmt = wb.CreateStyle();
            fmt.ForegroundColor = System.Drawing.Color.Yellow;
            fmt.Pattern = BackgroundType.Solid;
            cells["C1"].SetStyle(fmt);
            cells["C2"].SetStyle(fmt);
            cells["C3"].SetStyle(fmt);
            cells["D5"].SetStyle(fmt); // isolated formatting‑only cell

            // Additional formatting on a data cell (will stay after removal)
            Style dataStyle = wb.CreateStyle();
            dataStyle.Font.IsBold = true;
            cells["B2"].SetStyle(dataStyle);

            return wb;
        }
    }
}
