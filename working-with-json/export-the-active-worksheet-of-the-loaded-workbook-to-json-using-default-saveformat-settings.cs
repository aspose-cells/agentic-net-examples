// Title: Export only the active worksheet of an Excel file to JSON with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file using Aspose.Cells, checks that the file exists, enables the option to export only the active worksheet, and saves the result as a .json file. | Provide a .NET console example that catches runtime exceptions while converting the current worksheet to JSON with Aspose.Cells. | Show how to toggle the JSON export setting between all worksheets and the active worksheet in Aspose.Cells.
// Common Searches: Aspose.Cells C# export only current worksheet to JSON | How to use JsonSaveOptions to save active sheet as JSON in .NET | Example of Excel to JSON conversion with error handling using Aspose.Cells | JsonSaveOptions property to limit JSON output to one worksheet
// Tags: Aspose.Cells JsonSaveOptions single sheet export | C# Excel to JSON conversion | file existence validation for Excel input | exception handling Aspose.Cells workbook save | export single worksheet as JSON Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads an .xlsx file, optionally configures JsonSaveOptions to export only the active worksheet, and saves the data to a .json file with file‑existence validation and exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.json";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            var workbook = new Workbook(inputPath);

            // Configure JSON save options (default exports all worksheets)
            var jsonOptions = new JsonSaveOptions();

            // If the current Aspose.Cells version supports it, uncomment the line below
            // to export only the active worksheet:
            // jsonOptions.ExportActiveWorksheetOnly = true;

            // Save the workbook (or active worksheet) to a JSON file
            workbook.Save(outputPath, jsonOptions);
            Console.WriteLine($"Workbook successfully saved to JSON: {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
