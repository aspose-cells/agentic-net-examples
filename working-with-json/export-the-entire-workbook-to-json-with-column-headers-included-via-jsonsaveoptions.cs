// Title: Export a complete Excel workbook to JSON with column headers using Aspose.Cells JsonSaveOptions in C#
// Common Searches: C# Aspose.Cells export entire workbook to JSON including column names | How to save an Excel file as JSON with headers using Aspose.Cells JsonSaveOptions | Aspose.Cells JsonSaveOptions default behavior column headers C# example | Convert .xlsx to .json preserving column headings with Aspose.Cells | Save workbook to JSON format with column titles using Aspose.Cells library in .NET
// Tags: Aspose.Cells JsonSaveOptions workbook to JSON | C# Excel to JSON conversion with headers | serialize Excel workbook as JSON using Aspose | preserve column names when saving Excel as JSON .NET | Aspose.Cells export workbook JSON format

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Json;

// The sample checks for the existence of input.xlsx, loads it into an Aspose.Cells Workbook, applies the default JsonSaveOptions (which automatically includes column headers), and saves the entire workbook as output.json, handling any runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.json";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Configure JSON save options (default behavior includes column headers)
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Save the workbook as a JSON file
            workbook.Save(outputPath, jsonOptions);
            Console.WriteLine($"Workbook successfully saved to JSON: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
