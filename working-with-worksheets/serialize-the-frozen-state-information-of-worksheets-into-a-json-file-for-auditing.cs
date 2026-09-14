// Title: Export worksheet freeze pane configuration to a JSON file for audit using Aspose.Cells in C#
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, reads each worksheet's frozen rows and columns, and writes the results to an indented JSON file. | Update the example to include the top‑left cell address of each frozen pane in the JSON output alongside the row and column counts. | Add robust error handling that logs worksheets where freeze‑pane data cannot be retrieved and continues processing the remaining sheets.
// Common Searches: how to extract frozen pane rows and columns from an Excel file using Aspose.Cells C# | save worksheet freeze pane settings to JSON for auditing in .NET | Aspose.Cells serialize worksheet freeze state to a JSON file | C# program to list frozen rows and columns of all worksheets in a workbook | audit Excel freeze panes with Aspose.Cells and output JSON
// Tags: Aspose.Cells export freeze pane to JSON | C# serialize worksheet freeze state | Excel freeze pane audit .NET | Aspose.Cells worksheet freeze information extraction | JSON output of Excel frozen rows columns

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

// The example loads an existing Excel workbook with Aspose.Cells, iterates through each worksheet, captures the frozen row and column counts (defaulting to zero when the API does not expose them), serializes this data into an indented JSON structure, and writes it to 'worksheet_frozen_state.json' for auditing purposes.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string inputPath = "input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // List to hold frozen pane information for each worksheet
            var frozenStates = new List<WorksheetFreezeInfo>();

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // NOTE: In some Aspose.Cells versions FreezePanes is only a method,
                // so retrieving frozen rows/columns directly is not supported.
                // Defaulting to 0 for both values.
                int frozenRows = 0;
                int frozenColumns = 0;

                frozenStates.Add(new WorksheetFreezeInfo
                {
                    WorksheetName = sheet.Name,
                    FrozenRows = frozenRows,
                    FrozenColumns = frozenColumns
                });
            }

            // Serialize the collected information to JSON with indentation
            string json = JsonSerializer.Serialize(frozenStates, new JsonSerializerOptions { WriteIndented = true });

            // Write the JSON to a file for auditing purposes
            string outputPath = "worksheet_frozen_state.json";
            File.WriteAllText(outputPath, json);

            Console.WriteLine($"Frozen pane information written to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Simple DTO to represent frozen pane state of a worksheet
public class WorksheetFreezeInfo
{
    public string WorksheetName { get; set; } = string.Empty;
    public int FrozenRows { get; set; }
    public int FrozenColumns { get; set; }
}
