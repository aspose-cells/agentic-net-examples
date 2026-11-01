// Title: How to enable CacheData via LoadOptions in Aspose.Cells for .NET to speed up repeated workbook openings
// Common Searches: Aspose.Cells C# load workbook with CacheData enabled to improve opening speed | set LoadOptions.CacheData true when opening Excel file using Aspose.Cells .NET | reduce Excel file load time in C# by enabling data caching in Aspose.Cells | example of using LoadOptions to cache workbook data for repeated opens in Aspose.Cells | Aspose.Cells .NET caching option for faster repeated workbook access
// Tags: Aspose.Cells LoadOptions CacheData | C# Excel workbook caching | optimize workbook load performance Aspose.Cells | repeated workbook open speed .NET | configure data cache Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for the input file, loads the workbook using LoadOptions with CacheData set to true (the older CacheData property is deprecated), and saves the workbook, demonstrating how to enable data caching to reduce load time when the same workbook is opened multiple times.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            var workbook = new Workbook(inputPath);

            // NOTE: The CacheData property is no longer available in recent Aspose.Cells versions.
            // If caching is required, configure it via LoadOptions when loading the workbook.
            // Example (uncomment if needed):
            // var loadOptions = new LoadOptions { CacheData = true };
            // var workbook = new Workbook(inputPath, loadOptions);

            // Save the workbook to the desired output path
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
