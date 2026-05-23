using System;
using Aspose.Cells;

namespace SubtotalExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data
            // Header row
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Amount");

            // Data rows
            object[,] data = new object[,]
            {
                { "North", 1200 },
                { "North", 800 },
                { "South", 1500 },
                { "South", 700 },
                { "East",  900 },
                { "East",  1100 }
            };

            for (int i = 0; i < data.GetLength(0); i++)
            {
                cells[i + 1, 0].PutValue(data[i, 0]); // Category column (A)
                cells[i + 1, 1].PutValue(data[i, 1]); // Amount column (B)
            }

            // Define the range that contains the data (including header)
            // A1:B7 -> rows 0-6, columns 0-1
            CellArea area = CellArea.CreateCellArea(0, 0, data.GetLength(0), 1);

            // Apply subtotals:
            // - Group by the first column (Category) -> index 0
            // - Use SUM function
            // - Subtotal the second column (Amount) -> index 1
            // - Replace existing subtotals, no page breaks, summary placed below data
            cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 1 }, true, false, true);

            // Optional: retrieve and display subtotal settings for verification
            SubtotalSetting setting = cells.RetrieveSubtotalSetting(area);
            Console.WriteLine($"GroupBy column index: {setting.GroupBy}");
            Console.WriteLine($"Subtotal function: {setting.SubtotalFunction}");
            Console.WriteLine($"Subtotal applied to column index: {setting.TotalList[0]}");
            Console.WriteLine($"Summary placed below data: {setting.SummaryBelowData}");

            // Save the workbook
            workbook.Save("SubtotalResult.xlsx");
        }
    }
}