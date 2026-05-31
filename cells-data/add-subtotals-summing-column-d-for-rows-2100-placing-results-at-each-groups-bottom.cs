using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Sample data (rows 2‑100) – column D will be summed
                for (int row = 1; row < 100; row++) // zero‑based index: row 1 = Excel row 2
                {
                    // Group column A (replace with your own grouping logic)
                    cells[row, 0].PutValue(row % 5 == 0 ? "GroupB" : "GroupA");
                    // Column D (index 3) contains numeric values to be subtotaled
                    cells[row, 3].PutValue(row * 10);
                }

                // Define the range that includes rows 2‑100 and columns A‑D
                CellArea area = new CellArea
                {
                    StartRow = 1,      // Excel row 2
                    StartColumn = 0,   // Column A
                    EndRow = 99,       // Excel row 100
                    EndColumn = 3      // Column D
                };

                // Apply subtotals:
                // - Group by column A (index 0)
                // - Use SUM function
                // - Subtotal column D (index 3)
                // - Do not replace existing subtotals, no page breaks, place summary below each group
                cells.Subtotal(
                    area,
                    0,                                 // groupBy column index
                    ConsolidationFunction.Sum,         // subtotal function
                    new int[] { 3 },                   // totalList – column D
                    false,                             // replace existing subtotals
                    false,                             // add page breaks between groups
                    true                               // place summary below the data
                );

                // Save the workbook
                string outputPath = "SubtotalResult.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}