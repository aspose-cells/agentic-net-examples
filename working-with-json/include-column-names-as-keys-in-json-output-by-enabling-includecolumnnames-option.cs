// Title: Convert the first worksheet of an Excel workbook to a JSON file with column headers as property names using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells and saves the first worksheet to a JSON document where each column header becomes a JSON field. | Show how to configure JsonSaveOptions.IncludeColumnNames (or rely on its default) to ensure column headers are used as keys when exporting a worksheet to JSON with Aspose.Cells. | Add robust error handling that checks for the existence of the source Excel file and catches any exceptions during the conversion to JSON.
// Common Searches: aspnet convert first sheet of Excel to JSON using Aspose.Cells include column names | c# Aspose.Cells JsonSaveOptions include column headers as keys | how to export worksheet to JSON with header names in Aspose.Cells .NET | save Excel worksheet as JSON file preserving column titles Aspose.Cells
// Tags: Aspose.Cells export worksheet to JSON with headers | JsonSaveOptions IncludeColumnNames .NET | C# convert .xlsx to JSON Aspose.Cells | first worksheet JSON export Aspose.Cells | file existence check before Excel to JSON conversion

using Aspose.Cells;
using System;
using System.IO;

// // Loads an Excel workbook, configures JsonSaveOptions to include column names as JSON keys, and saves the first worksheet to a JSON file while handling missing input files and runtime errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.json";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure JSON save options (default behavior includes column names)
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Save the first worksheet to JSON using the configured options
            workbook.Save(outputPath, jsonOptions);

            Console.WriteLine($"Workbook successfully saved to JSON: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
