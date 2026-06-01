using System;
using Aspose.Cells;

namespace AsposeCellsWatchWindowDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // (Optional) Add some initial watches to demonstrate clearing later
            sheet.CellWatches.Add("A1");
            sheet.CellWatches.Add("B2");

            // Clear the Watch Window before adding a new set of cells
            sheet.CellWatches.Clear();

            // Add new cells to be monitored
            sheet.CellWatches.Add("C3");
            sheet.CellWatches.Add("D4");

            // Save the workbook to verify the watch window contains only the new watches
            workbook.Save("WatchWindowCleared.xlsx");
        }
    }
}