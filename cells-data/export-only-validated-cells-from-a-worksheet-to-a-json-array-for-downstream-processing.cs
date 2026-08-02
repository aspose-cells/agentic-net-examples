// Title: Export a Worksheet’s Used Range to JSON with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, selects the first worksheet, determines its used rows and columns, creates a range covering all populated cells, configures JsonSaveOptions to skip empty cells and rows while treating the first row as headers, and writes the resulting JSON array to a file for downstream processing.
// Keywords: Aspose.Cells | C# | .NET | Excel to JSON | JsonSaveOptions | JsonUtility ExportRangeToJson | skip empty rows | export used range | header row JSON | data extraction from Excel
// Common Searches: How to export Excel worksheet to JSON using Aspose.Cells C# | Aspose.Cells JsonSaveOptions skip empty cells | Export used range of a sheet to JSON file | C# convert Excel data to JSON with Aspose | Aspose.Cells ExportRangeToJson example
// Developer Intent: Create a JSON file that contains only the populated cells of the first worksheet, excluding blanks and preserving column headers.
// Use Cases: Feed Excel‑derived data into a REST API that expects a compact JSON payload. | Generate lightweight JSON reports for front‑end JavaScript applications without transmitting empty rows. | Prepare data for machine‑learning pipelines where only non‑empty, typed values are required.
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, selects the first worksheet’s used range, and exports it to JSON while skipping empty cells and rows. | Show how to configure JsonSaveOptions in Aspose.Cells to include a header row and export values as native types. | Explain error handling for missing input files and exceptions when converting Excel data to JSON with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Loads an Excel workbook, selects the first worksheet, determines its used rows and columns, creates a range covering all populated cells, configures JsonSaveOptions to skip empty cells and rows while treating the first row as headers, and writes the resulting JSON array to a file for downstream processing.
class ExportValidatedCellsToJson
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "validated_cells.json";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the used range (zero‑based indices)
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            // Create a range that covers all used cells
            Aspose.Cells.Range range = sheet.Cells.CreateRange(0, 0, maxRow + 1, maxCol + 1);

            // Configure JSON export options
            JsonSaveOptions options = new JsonSaveOptions
            {
                ExportEmptyCells = false,
                SkipEmptyRows = true,
                HasHeaderRow = true,
                ExportAsString = false
            };

            // Export the defined range to a JSON string
            string json = JsonUtility.ExportRangeToJson(range, options);

            // Output the JSON to console
            Console.WriteLine(json);

            // Write the JSON to a file
            File.WriteAllText(outputPath, json);
            Console.WriteLine($"JSON successfully written to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
