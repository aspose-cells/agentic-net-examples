using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalFilteredDataDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data with a header row
            // Columns: Region (A), Product (B), Sales (C)
            cells["A1"].PutValue("Region");
            cells["B1"].PutValue("Product");
            cells["C1"].PutValue("Sales");

            object[,] data = new object[,]
            {
                {"North", "Widget", 5000},
                {"North", "Gadget", 3000},
                {"South", "Widget", 6000},
                {"South", "Gadget", 4000},
                {"West",  "Widget", 4500},
                {"West",  "Gadget", 3500}
            };

            // Fill data starting from row 2 (zero‑based index 1)
            for (int i = 0; i < data.GetLength(0); i++)
            {
                cells[i + 1, 0].PutValue(data[i, 0]); // Region
                cells[i + 1, 1].PutValue(data[i, 1]); // Product
                cells[i + 1, 2].PutValue(data[i, 2]); // Sales
            }

            // Apply an AutoFilter on the header row (A1:C1)
            worksheet.AutoFilter.Range = "A1:C1";

            // Filter to show only rows where Region = "North"
            // Column index 0 corresponds to "Region"
            worksheet.AutoFilter.AddFilter(0, "North");
            // Refresh the filter to hide non‑matching rows
            worksheet.AutoFilter.Refresh();

            // Verify which rows are hidden (for demonstration)
            Console.WriteLine("Row visibility after filter:");
            for (int row = 1; row <= data.GetLength(0); row++)
            {
                bool hidden = cells.IsRowHidden(row);
                Console.WriteLine($"Row {row + 1} hidden: {hidden}");
            }

            // Define the cell area that includes the header and all data rows
            // StartRow = 0, StartColumn = 0 (A1), EndRow = data rows count + 1, EndColumn = 2 (C)
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = data.GetLength(0), // includes header row
                EndColumn = 2
            };

            // Add subtotals:
            // - Group by column 0 (Region)
            // - Use SUM function
            // - Subtotal the Sales column (index 2)
            // - Replace existing subtotals, no page breaks, summary placed below data
            cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 2 }, true, false, true);

            // Retrieve the subtotal setting to confirm configuration
            SubtotalSetting setting = cells.RetrieveSubtotalSetting(area);
            Console.WriteLine("\nSubtotal Setting:");
            Console.WriteLine($"GroupBy column index: {setting.GroupBy}");
            Console.WriteLine($"Function: {setting.SubtotalFunction}");
            Console.WriteLine($"SummaryBelowData: {setting.SummaryBelowData}");
            Console.WriteLine($"TotalList column indexes: {string.Join(",", setting.TotalList)}");

            // Compute the sum of visible Sales values manually to confirm that hidden rows are excluded
            double visibleSum = 0;
            for (int row = 1; row <= data.GetLength(0); row++)
            {
                if (!cells.IsRowHidden(row))
                {
                    visibleSum += Convert.ToDouble(cells[row, 2].Value);
                }
            }
            Console.WriteLine($"\nManual sum of visible Sales rows: {visibleSum}");

            // Save the workbook
            workbook.Save("SubtotalFilteredDataDemo.xlsx");
        }
    }
}