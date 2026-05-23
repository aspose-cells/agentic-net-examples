using System;
using Aspose.Cells;

namespace AsposeCellsWatchWindowReport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (or load an existing one if needed)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample cell watches for demonstration
            sheet.CellWatches.Add("B2");
            sheet.CellWatches.Add("C3");
            sheet.CellWatches.Add("D4");

            // Retrieve the collection of watched cells
            CellWatchCollection watches = sheet.CellWatches;

            // Report each watched cell's details
            Console.WriteLine("Watched Cells:");
            foreach (CellWatch watch in watches)
            {
                // Row and Column are zero‑based indices
                Console.WriteLine($"Name: {watch.CellName}, Row: {watch.Row}, Column: {watch.Column}");
            }

            // Save the workbook (optional, just to keep the watch window data)
            workbook.Save("WatchWindowReport.xlsx");
        }
    }
}