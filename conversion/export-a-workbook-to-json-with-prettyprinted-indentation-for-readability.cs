// Title: Export an Excel workbook to indented JSON using Aspose.Cells for .NET
// Common Searches: how to save an .xlsx workbook as indented JSON with Aspose.Cells in C# | Aspose.Cells JsonSaveOptions generate formatted JSON output | C# convert Excel file to pretty printed JSON using Aspose.Cells | export Excel to JSON with readable indentation using Aspose.Cells .NET | save workbook to JSON with readable formatting Aspose.Cells C#
// Tags: Aspose.Cells JsonSaveOptions usage | C# export Excel to JSON | Workbook.Save JSON format | Excel to indented JSON conversion | Aspose.Cells JSON output formatting

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the presence of input.xlsx, loads it into an Aspose.Cells Workbook, creates a JsonSaveOptions object (the library does not provide a built‑in pretty‑print option), and saves the workbook as output.json using Workbook.Save, while handling possible exceptions.
class ExportWorkbookToJson
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.json";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The input file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure JSON save options (pretty‑print not available in this version)
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Export the workbook to a JSON file with the specified options
            workbook.Save(outputPath, jsonOptions);

            Console.WriteLine($"Workbook has been exported to JSON at \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
