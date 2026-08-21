// Title: C# – List Cells Monitored by the Watch Window with Aspose.Cells
// Description: Shows how to add CellWatch entries (A1, B2, C3) to a worksheet, iterate the CellWatches collection, output each watch’s address, row and column, and save the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | CellWatches | watch window | C# example | enumerate watched cells | CellWatch.Row | CellWatch.Column | debug workbook | retrieve cell watches
// Common Searches: Aspose.Cells get all cell watches | list watched cells in worksheet C# | enumerate CellWatch collection Aspose | how to read watch window cells Aspose.Cells
// Developer Intent: Obtain the full collection of cells currently tracked by the Watch Window in a worksheet.
// Use Cases: Create a debugging report that lists every watched cell with its address, row, and column. | Log watched cell addresses for audit before exporting the workbook. | Validate that required cells are being monitored by checking the CellWatches collection programmatically.
// AI Prompts: Generate C# code to filter CellWatches by a specific row range in Aspose.Cells. | Provide an example that removes a particular watch (e.g., B2) from a worksheet. | Explain how to serialize the CellWatch collection to JSON for external reporting.

using System;
using Aspose.Cells;

// Shows how to add CellWatch entries (A1, B2, C3) to a worksheet, iterate the CellWatches collection, output each watch’s address, row and column, and save the workbook using Aspose.Cells for .NET.
class RetrieveCellWatches
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add some cell watches for demonstration
        sheet.CellWatches.Add("A1");
        sheet.CellWatches.Add("B2");
        sheet.CellWatches.Add("C3");

        // Retrieve and list all cells currently monitored by the Watch Window
        Console.WriteLine("Watched Cells:");
        foreach (CellWatch watch in sheet.CellWatches)
        {
            Console.WriteLine($"- {watch.CellName} (Row: {watch.Row}, Column: {watch.Column})");
        }

        // Save the workbook (lifecycle rule)
        workbook.Save("WatchedCellsDemo.xlsx");
    }
}
