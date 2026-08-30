// Title: C# example to clear the Aspose.Cells watch window and add new cell watches to a worksheet
// AI Prompts: Generate C# code that removes all current CellWatches from a worksheet using Aspose.Cells, then adds watches for cells A1 and D3, writes values, and saves the workbook. | Show how to programmatically reset the watch window in an Aspose.Cells workbook and monitor specific cells in a .NET application.
// Common Searches: asp.net aspose.cells clear watch window before adding cell watches | c# remove existing cell watches from Aspose.Cells workbook | how to add cell watches programmatically after resetting watch window in Aspose.Cells
// Tags: clear cell watches Aspose.Cells | add cell watches C# | watch window reset .NET | cell monitoring Aspose.Cells workbook | Aspose.Cells CellWatches API

using System;
using Aspose.Cells;

namespace AsposeCellsWatchWindowDemo
{
    // Demonstrates how to clear the worksheet's CellWatches collection, add watches for cells A1 and D3, assign values to those cells, and save the workbook as an .xlsx file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Clear any existing cell watches in the watch window
            sheet.CellWatches.Clear();

            // Add new cells to monitor
            sheet.CellWatches.Add("A1");          // Watch cell A1
            sheet.CellWatches.Add(2, 3);          // Watch cell D3 (row 2, column 3, zero‑based)

            // Optionally put some values in the watched cells
            sheet.Cells["A1"].PutValue("Watch me");
            sheet.Cells[2, 3].PutValue("Also watch me");

            // Save the workbook
            workbook.Save("WatchWindowDemo.xlsx");
        }
    }
}
