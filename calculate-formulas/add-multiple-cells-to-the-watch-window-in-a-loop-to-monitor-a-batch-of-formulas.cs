using System;
using Aspose.Cells;

namespace AsposeCellsWatchWindowDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (using the standard creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample formulas in column A (A1:A10)
            for (int i = 0; i < 10; i++)
            {
                // Example formula: each cell sums its row number with a constant
                sheet.Cells[i, 0].Formula = $"=ROW()+5";
            }

            // Get the CellWatchCollection for the worksheet
            CellWatchCollection watches = sheet.CellWatches;

            // Add watches for the range A1:A10 using a loop
            for (int row = 0; row < 10; row++)
            {
                // Convert row index to cell name (e.g., A1, A2, ...)
                string cellName = CellsHelper.CellIndexToName(row, 0);
                watches.Add(cellName);
            }

            // Display the total number of watches added
            Console.WriteLine($"Total Cell Watches added: {watches.Count}");

            // Optionally, iterate and print each watched cell name
            foreach (CellWatch watch in watches)
            {
                Console.WriteLine($"Watching cell: {watch.CellName}");
            }

            // Save the workbook (using the standard save rule)
            workbook.Save("WatchWindowDemo.xlsx");
        }
    }
}