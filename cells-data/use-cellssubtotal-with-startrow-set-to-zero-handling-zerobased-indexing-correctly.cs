using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data (including header row at index 0)
            // Header
            cells[0, 0].PutValue("Region");
            cells[0, 1].PutValue("Product");
            cells[0, 2].PutValue("Sales");

            // Data rows (zero‑based indexing, rows 1‑5)
            object[,] data = new object[,]
            {
                { "North", "Widget", 5000 },
                { "North", "Gadget", 3000 },
                { "South", "Widget", 6000 },
                { "South", "Gadget", 4000 },
                { "West",  "Widget", 4500 }
            };

            for (int r = 0; r < data.GetLength(0); r++)
            {
                for (int c = 0; c < data.GetLength(1); c++)
                {
                    // Row index is offset by 1 because row 0 holds headers
                    cells[r + 1, c].PutValue(data[r, c]);
                }
            }

            // Define the cell area that includes the header row (A1:C6)
            // StartRow = 0 ensures zero‑based indexing is handled correctly
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 5,   // rows 0‑5 inclusive (6 rows total)
                EndColumn = 2 // columns A‑C
            };

            // Apply subtotal:
            // - Group by column 0 (Region)
            // - Use SUM function
            // - Add subtotal for column 2 (Sales)
            cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 2 });

            // Optional: retrieve and display the subtotal settings
            SubtotalSetting setting = cells.RetrieveSubtotalSetting(area);
            Console.WriteLine($"GroupBy index: {setting.GroupBy}");
            Console.WriteLine($"Subtotal function: {setting.SubtotalFunction}");
            Console.WriteLine($"Total columns: {string.Join(",", setting.TotalList)}");
            Console.WriteLine($"Summary below data: {setting.SummaryBelowData}");

            // Save the workbook
            workbook.Save("SubtotalDemo.xlsx");
        }
    }
}