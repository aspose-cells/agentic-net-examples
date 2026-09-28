// Title: How to list cells monitored by the Watch Window using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, accesses the worksheet's CellWatches collection, and prints each watched cell's address, row, and column. | Create a method that returns a formatted string containing all watch‑window entries (cell name, row, column) from a specified worksheet using Aspose.Cells. | Demonstrate how to iterate over Worksheet.CellWatches in Aspose.Cells and output the watch details to the console or a log file.
// Common Searches: Aspose.Cells C# retrieve watch window cell list from worksheet | How to enumerate CellWatchCollection in Aspose.Cells .NET | Get monitored cells from an Excel file using Aspose.Cells API | List watch window entries programmatically with Aspose.Cells for .NET | Extract cell watch information from a workbook using C# Aspose.Cells
// Tags: Aspose.Cells enumerate CellWatchCollection | C# list watch window cells | Aspose.Cells retrieve monitored cells .xlsx | Worksheet.CellWatches iteration | Aspose.Cells watch window reporting

using System;
using Aspose.Cells;

namespace AsposeCellsWatchWindowReport
{
    // The example loads an Excel workbook, accesses the first worksheet, obtains its CellWatches collection, and writes each watched cell's name, row, and column to the console, illustrating how to report watch‑window entries with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Get the first worksheet (or any specific worksheet)
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the collection of cell watches for this worksheet
            CellWatchCollection watches = sheet.CellWatches;

            // Report each watched cell's details
            Console.WriteLine("Watched Cells:");
            foreach (CellWatch watch in watches)
            {
                Console.WriteLine($"Cell Name: {watch.CellName}, Row: {watch.Row}, Column: {watch.Column}");
            }

            // Optionally, save the workbook if any changes were made
            // workbook.Save("output.xlsx");
        }
    }
}
