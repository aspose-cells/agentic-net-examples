using System;
using Aspose.Cells;

class VerifyCellWatches
{
    static void Main()
    {
        // Path to the workbook that already contains cell watches.
        // The workbook should have been created and saved earlier using the
        // Worksheet.CellWatches.Add method (e.g., "CellWatchDemo.xlsx").
        string filePath = "CellWatchDemo.xlsx";

        // Load the workbook from the file system.
        // This uses the Workbook(string) constructor rule.
        Workbook workbook = new Workbook(filePath);

        // Access the first worksheet in the workbook.
        Worksheet sheet = workbook.Worksheets[0];

        // Cells that we expect to be present in the Watch Window.
        string[] expectedWatches = { "B2", "A1" };

        // Iterate through the expected cells and verify their presence.
        foreach (string cellName in expectedWatches)
        {
            // Retrieve the CellWatch object by cell name using the indexer.
            // If the cell is not being watched, the indexer returns null.
            CellWatch watch = sheet.CellWatches[cellName];

            if (watch != null)
            {
                Console.WriteLine($"Watch found: {watch.CellName} (Row={watch.Row}, Column={watch.Column})");
            }
            else
            {
                Console.WriteLine($"Watch NOT found for cell: {cellName}");
            }
        }

        // Output the total number of watches in the worksheet.
        Console.WriteLine($"Total watches in worksheet: {sheet.CellWatches.Count}");
    }
}