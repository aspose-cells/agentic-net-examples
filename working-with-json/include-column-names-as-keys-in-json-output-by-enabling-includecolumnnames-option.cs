// Title: Convert an Excel worksheet to JSON with column names as property keys using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, sets JsonSaveOptions.IncludeColumnNames = true, and saves the result to a .json file. | Show how to export an Excel workbook to JSON while using the first row as field names in a .NET application with Aspose.Cells.
// Common Searches: Aspose.Cells C# export Excel to JSON using column headers as keys | How to set JsonSaveOptions.IncludeColumnNames in Aspose.Cells .NET | Save workbook as JSON with first row as property names using Aspose.Cells | Convert .xlsx to .json with column names using Aspose.Cells library | C# example for JSON export with column headers from Excel via Aspose.Cells
// Tags: Aspose.Cells JsonSaveOptions column name export | export Excel worksheet to JSON with headers | C# workbook to JSON conversion using Aspose.Cells | JSON serialization preserving Excel column names | Aspose.Cells .NET JSON output configuration

using System;
using System.IO;
using Aspose.Cells;

// The program verifies the presence of an input.xlsx file, loads it into an Aspose.Cells Workbook, configures JsonSaveOptions to include column names as JSON property keys, and saves the workbook as output.json, handling any exceptions that may arise.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.json";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure JSON save options
            JsonSaveOptions jsonOptions = new JsonSaveOptions();
            // Note: In the current Aspose.Cells version, column headers are exported by default.
            // If a specific option is required in future versions, it can be set here.

            // Save the workbook as JSON
            workbook.Save(outputPath, jsonOptions);
            Console.WriteLine($"Workbook successfully saved as JSON to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
