// Title: Calculate total merged cell regions in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# count merged cells across all worksheets | how to get number of merged regions in an Excel file with Aspose.Cells | C# enumerate worksheets and sum merged cell ranges using Aspose.Cells | retrieve total merged cell count from workbook with Aspose.Cells .NET | Aspose.Cells merged cells collection count example
// Tags: Aspose.Cells merged cells count | C# total merged regions workbook | enumerate worksheets merged cells Aspose.Cells | merged cells collection Aspose.Cells .NET | calculate merged cell regions Excel Aspose

using Aspose.Cells;
using System;
using System.IO;

// The sample checks for the input.xlsx file, loads it into an Aspose.Cells Workbook, iterates each worksheet, accesses the Cells.MergedCells collection, adds each sheet's merged region count to a total, and prints the overall number of merged cell regions while handling Aspose.Cells‑specific and generic exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: File not found – \"{inputPath}\"");
            return;
        }

        try
        {
            // Load the workbook (no password supplied; adjust LoadOptions if needed)
            var loadOptions = new LoadOptions(); // Add Password = "yourPassword" if required
            Workbook workbook = new Workbook(inputPath, loadOptions);

            int totalMergedRegions = 0;

            // Enumerate each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the collection of merged cells for the current worksheet
                var mergedCells = sheet.Cells.MergedCells;

                // Add the count of merged regions in this worksheet to the total
                totalMergedRegions += mergedCells.Count;
            }

            // Output the total number of merged regions across all worksheets
            Console.WriteLine("Total merged regions: " + totalMergedRegions);
        }
        catch (CellsException ex)
        {
            // Handle Aspose.Cells specific errors (e.g., password protection)
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handle any other unexpected errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
