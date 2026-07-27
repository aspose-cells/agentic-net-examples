using System;
using Aspose.Cells;

namespace AsposeCellsWatchDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add watches for a batch of cells (A1 to A10) in a loop
            for (int i = 0; i < 10; i++)
            {
                // Define the cell address (e.g., A1, A2, ...)
                string cellName = $"A{i + 1}";

                // Optionally set a formula in the cell to be monitored
                sheet.Cells[cellName].Formula = $"=ROW()*2";

                // Add the cell to the Watch Window
                sheet.CellWatches.Add(cellName);
            }

            // Output the total number of watches added
            Console.WriteLine($"Total watches added: {sheet.CellWatches.Count}");

            // Save the workbook with the watch window configuration
            workbook.Save("WatchWindowDemo.xlsx");
        }
    }
}