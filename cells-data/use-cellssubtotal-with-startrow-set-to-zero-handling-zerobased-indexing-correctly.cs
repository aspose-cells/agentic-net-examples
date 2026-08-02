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

            // Populate sample data (including header row at row 0)
            // Header: Category | Product | Amount
            cells[0, 0].PutValue("Category");
            cells[0, 1].PutValue("Product");
            cells[0, 2].PutValue("Amount");

            object[,] data = new object[,]
            {
                { "North", "Widget", 5000 },
                { "North", "Gadget", 3000 },
                { "South", "Widget", 6000 },
                { "South", "Gadget", 4000 },
                { "West",  "Widget", 4500 }
            };

            // Fill data starting from row index 1 (zero‑based)
            for (int i = 0; i < data.GetLength(0); i++)
            {
                cells[i + 1, 0].PutValue(data[i, 0]); // Category
                cells[i + 1, 1].PutValue(data[i, 1]); // Product
                cells[i + 1, 2].PutValue(data[i, 2]); // Amount
            }

            // Define the range that includes the header row (A1:C6)
            // StartRow = 0 because we want to include the header row in the subtotal range
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = data.GetLength(0), // last data row index (5)
                EndColumn = 2
            };

            // Apply subtotal:
            // groupBy = 0 (Category column, zero‑based)
            // function = Sum
            // totalList = new int[] { 2 } (Amount column, zero‑based)
            // replace = true, pageBreaks = false, summaryBelowData = true
            cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 2 }, true, false, true);

            // Optional: retrieve the subtotal setting to verify parameters
            SubtotalSetting setting = cells.RetrieveSubtotalSetting(area);
            Console.WriteLine($"GroupBy column index: {setting.GroupBy}");
            Console.WriteLine($"Subtotal function: {setting.SubtotalFunction}");
            Console.WriteLine($"Summary below data: {setting.SummaryBelowData}");
            Console.WriteLine($"Total columns: {string.Join(",", setting.TotalList)}");

            // Save the workbook
            workbook.Save("SubtotalWithZeroBasedStartRow.xlsx");
        }
    }
}