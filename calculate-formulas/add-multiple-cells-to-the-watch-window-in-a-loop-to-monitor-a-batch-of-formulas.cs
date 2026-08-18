// Title: Loop to Add Multiple Cells to the Watch Window with Aspose.Cells for .NET (C#)
// Description: This C# example creates a workbook, writes simple formulas to cells A1‑A10, then uses a loop to add each of those cells to the worksheet’s CellWatchCollection (watch window). It prints the total watched cells and saves the file as WatchWindowDemo.xlsx.
// Keywords: Aspose.Cells watch window | CellWatchCollection C# | add cells programmatically | monitor formulas .NET | batch watch cells | loop add watch cells | debug formulas Aspose | C# Excel watch window | Aspose.Cells API CellWatch | Excel formula evaluation monitoring
// Common Searches: add range to watch window Aspose.Cells C# | CellWatchCollection loop example | how to monitor multiple formulas with Aspose.Cells | watch window batch cells .NET | debug Excel formulas using Aspose.Cells watch window
// Developer Intent: Programmatically add a set of formula cells to the worksheet’s watch window.
// Use Cases: Step‑by‑step debugging of a column of dependent formulas | Automatically track results of dynamically generated formulas during runtime | Generate a report of watched cells count before exporting the workbook | Integrate watch‑window population into a user‑driven range selection UI | Validate formula correctness in automated tests by observing watch window values
// AI Prompts: Write C# code to add cells B2:B20 to a CellWatchCollection using Aspose.Cells. | Show how to remove a specific cell from the watch window after evaluation. | Create a loop that adds all cells from a named range to the watch window and logs each addition. | Demonstrate how to clear the watch window before adding a new batch of cells. | Provide an example that combines adding watches with retrieving their evaluated values.

using System;
using Aspose.Cells;

namespace AsposeCellsWatchWindowDemo
{
    // This C# example creates a workbook, writes simple formulas to cells A1‑A10, then uses a loop to add each of those cells to the worksheet’s CellWatchCollection (watch window). It prints the total watched cells and saves the file as WatchWindowDemo.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample formulas in column A (A1:A10)
            for (int i = 0; i < 10; i++)
            {
                // Example formula: each cell sums its row number with a constant
                sheet.Cells[i, 0].Formula = $"=ROW() + 5";
            }

            // Get the CellWatchCollection for the worksheet
            CellWatchCollection watches = sheet.CellWatches;

            // Loop through the range A1:A10 and add each cell to the watch window
            for (int row = 0; row < 10; row++)
            {
                // Build the cell name (e.g., "A1", "A2", ...)
                string cellName = CellsHelper.CellIndexToName(row, 0);
                // Add the cell to the watch collection
                watches.Add(cellName);
            }

            // Optionally, display the count of watched cells
            Console.WriteLine($"Total cells added to watch window: {watches.Count}");

            // Save the workbook to a file
            workbook.Save("WatchWindowDemo.xlsx");
        }
    }
}
