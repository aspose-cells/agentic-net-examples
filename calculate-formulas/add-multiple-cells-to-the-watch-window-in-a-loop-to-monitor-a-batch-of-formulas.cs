// Title: Add multiple cells to the Aspose.Cells watch window in a C# loop to monitor a range of formulas
// AI Prompts: Generate C# code that iterates over a specified cell range and adds each cell to the worksheet's CellWatchCollection. | Write a C# snippet that logs the cell name and watch index for every added watch, then saves the workbook preserving the watch window. | Create a C# example that retrieves and prints the total number of watches after populating the watch window with a batch of formulas.
// Common Searches: c# Aspose.Cells add cell watches in a loop for range A1:A10 | how to programmatically monitor multiple formulas with Aspose.Cells watch window | saving watch window data to an Excel file using Aspose.Cells .NET | retrieve count of cell watches after adding them in Aspose.Cells C#
// Tags: add cells to CellWatchCollection Aspose.Cells | loop through range to create cell watches C# | monitor formula evaluation with watch window Aspose.Cells | save workbook with watch window data Aspose.Cells | retrieve watch count Aspose.Cells API

using System;
using Aspose.Cells;

namespace AsposeCellsWatchWindowDemo
{
    // The example creates a new workbook, fills cells A1:A10 with a simple formula, loops through the range to add each cell to the worksheet's CellWatchCollection, logs each cell name and watch index, displays the total watch count, and saves the workbook so the watch window information is retained.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample formulas in column A (A1:A10)
            for (int i = 0; i < 10; i++)
            {
                // Example formula: =ROW()*2
                sheet.Cells[i, 0].Formula = $"=ROW()*2";
            }

            // Add each of the cells in the range A1:A10 to the Watch Window
            CellWatchCollection watches = sheet.CellWatches;
            for (int row = 0; row < 10; row++)
            {
                // Add by row and column indices (0‑based). Column 0 corresponds to column A.
                int index = watches.Add(row, 0);
                // Optionally, retrieve the added watch to verify
                CellWatch watch = watches[index];
                Console.WriteLine($"Watch added for cell {watch.CellName} at index {index}");
            }

            // Display total number of watches added
            Console.WriteLine($"Total watches in the watch window: {watches.Count}");

            // Save the workbook (the watch window information is stored in the file)
            workbook.Save("WatchWindowDemo.xlsx");
        }
    }
}
