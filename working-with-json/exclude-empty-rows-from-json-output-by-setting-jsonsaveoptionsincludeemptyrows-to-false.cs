// Title: How to exclude empty rows when saving an Excel workbook to JSON using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# JsonSaveOptions IncludeEmptyRows false usage example | export Excel to JSON without blank rows using Aspose.Cells .NET | prevent empty rows in JSON output from Excel workbook in C# | set IncludeEmptyRows to false for JSON conversion with Aspose.Cells | C# Aspose.Cells JSON export omit empty rows tutorial
// Tags: Aspose.Cells JsonSaveOptions IncludeEmptyRows false | C# Excel to JSON conversion without empty rows | Aspose.Cells JSON export configuration | exclude blank rows JSON Aspose.Cells | Excel workbook JSON serialization Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, configures JsonSaveOptions with IncludeEmptyRows set to false, ensures the output directory exists, and saves the workbook as a JSON file, thereby omitting any empty rows from the generated JSON.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Configure JSON save options (default options are sufficient for most cases)
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Define the output path and ensure its directory exists
            string outputPath = "output.json";
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as JSON using the JsonSaveOptions
            workbook.Save(outputPath, jsonOptions);

            Console.WriteLine($"Workbook successfully saved as JSON to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
