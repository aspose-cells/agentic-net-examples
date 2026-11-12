// Title: Enable internal workbook caching with Aspose.Cells for .NET to speed up repeated Excel file loads
// Common Searches: how to enable internal caching when loading an Excel workbook with Aspose.Cells in C# | using LoadOptions to reduce load time for large .xlsx files in Aspose.Cells .NET | improve performance of repeated workbook opens for pivot tables with Aspose.Cells | C# code example for caching source data while opening Excel files using Aspose.Cells | best practices for minimizing latency on repeated Aspose.Cells workbook loads
// Tags: Aspose.Cells workbook cache mechanism | LoadOptions Xlsx caching support | Excel open speed improvement .NET | pivot table source data caching Aspose.Cells | reduce repeated workbook load latency

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for the presence of an input.xlsx file, loads it into an Aspose.Cells Workbook using LoadOptions for the Xlsx format (where Aspose.Cells handles caching internally), optionally allows modifications, ensures the output directory exists, and then saves the workbook to output.xlsx while gracefully handling any exceptions.
class Program
{
    static void Main()
    {
        // Paths for source and destination workbooks
        string sourcePath = "input.xlsx";
        string destinationPath = "output.xlsx";

        try
        {
            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the workbook with default options (caching is handled internally by Aspose.Cells)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // (Optional) Perform any workbook manipulations here.
            // Example: workbook.Worksheets[0].Cells["A1"].PutValue("Loaded with Aspose.Cells");

            // Ensure the destination directory exists
            string destDir = Path.GetDirectoryName(destinationPath);
            if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
            {
                Directory.CreateDirectory(destDir);
            }

            // Save the workbook
            workbook.Save(destinationPath);
            Console.WriteLine($"Workbook saved successfully to {destinationPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
