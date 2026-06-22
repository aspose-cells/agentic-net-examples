using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class FormattingOnlyCellsReductionDemo
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
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // -------------------------------------------------
            // 1. Populate cells with values and formatting‑only cells
            // -------------------------------------------------
            // Cells with actual data
            cells["A1"].PutValue("Data 1");
            cells["B2"].PutValue(123);
            cells["C3"].PutValue(DateTime.Now);

            // Cells that have only formatting (no value)
            Style fmtStyle = workbook.CreateStyle();
            fmtStyle.Font.Name = "Arial";
            fmtStyle.Font.Size = 14;
            fmtStyle.Font.IsBold = true;

            // Apply the style to cells without setting a value
            cells["D4"].SetStyle(fmtStyle);
            cells["E5"].SetStyle(fmtStyle);
            cells["F6"].SetStyle(fmtStyle);

            // -------------------------------------------------
            // 2. Count instantiated cells before removal
            // -------------------------------------------------
            long countBefore = cells.CountLarge;
            Console.WriteLine($"Instantiated cells before removal: {countBefore}");

            // -------------------------------------------------
            // 3. Remove formatting‑only cells
            //    A cell is considered formatting‑only if it has no value
            //    and its style differs from the default workbook style.
            // -------------------------------------------------
            Style defaultStyle = workbook.DefaultStyle;

            // Determine used range
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Skip cells that were never instantiated
                    if (cell == null) continue;

                    // Check if the cell has no value
                    bool hasNoValue = cell.Value == null || string.IsNullOrEmpty(cell.StringValue);

                    // Check if the cell's style is not the default one
                    bool hasCustomStyle = !cell.GetStyle().Equals(defaultStyle);

                    if (hasNoValue && hasCustomStyle)
                    {
                        // Reset style to default; this will de‑instantiate the cell
                        cell.SetStyle(defaultStyle);
                    }
                }
            }

            // -------------------------------------------------
            // 4. Count instantiated cells after removal
            // -------------------------------------------------
            long countAfter = cells.CountLarge;
            Console.WriteLine($"Instantiated cells after removal: {countAfter}");

            // -------------------------------------------------
            // 5. Show reduction
            // -------------------------------------------------
            long reduction = countBefore - countAfter;
            Console.WriteLine($"Number of cells removed: {reduction}");

            // Save the workbook for verification (optional)
            string outputPath = "FormattingOnlyCellsReductionDemo.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}