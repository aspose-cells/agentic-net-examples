// Title: Freeze rows in an Excel sheet with Aspose.Cells .NET using a JSON configuration
// Description: A C# console app that loads a `config.json` file, reads the `FreezeRows` value, creates a workbook, fills sample data, applies `Worksheet.FreezePanes` for the specified number of rows at runtime, and saves the result as `output.xlsx`.
// Keywords: Aspose.Cells | C# | .NET | FreezePanes | JSON config | dynamic freeze rows | Excel freeze panes runtime | read settings from file | worksheet freeze rows
// Common Searches: Aspose.Cells freeze rows from JSON file | C# set FreezePanes using configuration | read Excel freeze settings from config.json | dynamic FreezePanes example Aspose.Cells | how to freeze header rows at runtime in .NET
// Developer Intent: Load a JSON file to obtain the number of rows to freeze, then call `Worksheet.FreezePanes` with that value while generating the workbook.
// Use Cases: Allow non‑technical users to adjust header height by editing a simple config file. | Generate reports where each execution may require a different number of frozen rows. | Support multi‑tenant exports where each tenant’s settings dictate the freeze pane layout.
// AI Prompts: Write C# code that reads a `FreezeRows` setting from an XML file and applies `Worksheet.FreezePanes` with Aspose.Cells. | Show how to validate the `FreezeRows` value from `config.json` before calling `FreezePanes` to avoid out‑of‑range errors. | Demonstrate updating the frozen rows count after the workbook is created without rebuilding the entire file.

using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// A C# console app that loads a `config.json` file, reads the `FreezeRows` value, creates a workbook, fills sample data, applies `Worksheet.FreezePanes` for the specified number of rows at runtime, and saves the result as `output.xlsx`.
class Program
{
    static void Main()
    {
        // Load configuration from a JSON file (e.g., config.json)
        const string configPath = "config.json";
        if (!File.Exists(configPath))
        {
            Console.WriteLine($"Configuration file '{configPath}' not found.");
            return;
        }

        string json = File.ReadAllText(configPath);
        Config? cfg = JsonSerializer.Deserialize<Config>(json);
        int freezeRows = cfg?.FreezeRows ?? 0; // Default to 0 if not specified

        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add some sample data so the effect of freezing can be seen
        for (int i = 0; i < 20; i++)
        {
            sheet.Cells[i, 0].PutValue($"Row {i + 1}");
            sheet.Cells[i, 1].PutValue($"Data {i + 1}");
        }

        // Apply freeze panes based on the configuration (runtime operation)
        // Freeze 'freezeRows' rows from the top, no frozen columns
        if (freezeRows > 0)
        {
            // Freeze at the cell just below the frozen rows (row index = freezeRows, column = 0)
            // Parameters: row, column, freezedRows, freezedColumns
            sheet.FreezePanes(freezeRows, 0, freezeRows, 0);
        }

        // Save the workbook (lifecycle: save)
        const string outputPath = "output.xlsx";
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to '{outputPath}' with {freezeRows} frozen rows.");
    }

    // Simple POCO to map the JSON configuration
    private class Config
    {
        public int FreezeRows { get; set; }
    }
}
