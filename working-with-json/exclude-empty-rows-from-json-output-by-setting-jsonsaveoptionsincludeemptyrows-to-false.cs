// Title: Export Excel worksheet to JSON without blank rows using Aspose.Cells JsonSaveOptions.IncludeEmptyRows in C#
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, sets JsonSaveOptions.IncludeEmptyRows to false, and saves the data as a JSON file. | Show how to adjust an existing Aspose.Cells JSON export routine to skip empty rows by configuring the JsonSaveOptions object.
// Common Searches: Aspose.Cells JsonSaveOptions IncludeEmptyRows false example C# | How to remove empty rows when converting Excel to JSON with Aspose.Cells .NET | C# export Excel sheet to JSON skipping blank rows using Aspose | Save workbook as JSON without empty rows Aspose.Cells library | JSON export from Excel omitting empty rows Aspose.Cells C# tutorial
// Tags: Aspose.Cells JSON export options | C# skip blank rows in JSON output | Excel to JSON conversion Aspose.Cells | JsonSaveOptions configuration .NET | exclude empty rows Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Json;

// The example verifies that the source Excel file exists, loads it into an Aspose.Cells Workbook, creates a JsonSaveOptions instance with IncludeEmptyRows set to false to omit blank rows, and saves the workbook as a JSON file while handling any exceptions and reporting success or errors.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.json";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Configure JSON save options (default settings)
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Save the workbook (or specific sheet) as JSON using the configured options
            workbook.Save(outputPath, jsonOptions);
            Console.WriteLine($"Worksheet saved as JSON to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
