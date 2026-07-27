using System;
using Aspose.Cells;

namespace AsposeCellsWatchWindowDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Clear any existing cell watches (watch window)
            sheet.CellWatches.Clear();

            // Add new cells to monitor
            sheet.CellWatches.Add("A1");
            sheet.CellWatches.Add("B2");
            sheet.CellWatches.Add(2, 3); // adds cell at row 3, column D (0‑based)

            // Optionally, verify the count
            Console.WriteLine("Number of cell watches after clearing and adding: " + sheet.CellWatches.Count);

            // Save the workbook
            workbook.Save("WatchWindowClearedAndUpdated.xlsx");
        }
    }
}