using System;
using Aspose.Cells;

namespace WatchWindowReport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one if needed)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Example: add some watches for demonstration purposes
            sheet.CellWatches.Add("B2");
            sheet.CellWatches.Add("C3");
            sheet.CellWatches.Add("D4");

            // Retrieve and report all cells being watched
            Console.WriteLine("Cells monitored by the Watch Window:");
            foreach (CellWatch watch in sheet.CellWatches)
            {
                // Row and Column are zero‑based indices
                Console.WriteLine($"Name: {watch.CellName}, Row: {watch.Row}, Column: {watch.Column}");
            }

            // Save the workbook (optional, just to demonstrate lifecycle usage)
            workbook.Save("WatchWindowReport.xlsx");
        }
    }
}