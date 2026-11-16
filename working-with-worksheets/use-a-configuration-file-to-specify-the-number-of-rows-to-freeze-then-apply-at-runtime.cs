// Title: Read row freeze count from a JSON config and apply FreezePanes in Aspose.Cells for .NET
// Common Searches: how to read freeze rows value from a JSON file and set FreezePanes with Aspose.Cells C# | Aspose.Cells freeze top rows based on external configuration in .NET | C# load JSON settings to freeze rows in an Excel worksheet using Aspose.Cells | dynamic FreezePanes in Aspose.Cells using config.json | apply row freeze panes at runtime with Aspose.Cells and a JSON config
// Tags: JSON-driven FreezePanes with Aspose.Cells | dynamic row freeze in Excel using Aspose.Cells | read freeze rows from config file C# | apply FreezePanes at runtime Aspose.Cells | configurable worksheet freeze panes .NET

using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// The example reads a JSON configuration file (config.json) to obtain a "FreezeRows" value, loads an input.xlsx workbook with Aspose.Cells, and if the value is greater than zero, calls Worksheet.FreezePanes to freeze that number of rows before saving the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        // Path to the configuration file (JSON format)
        // Example content: { "FreezeRows": 5 }
        const string configPath = "config.json";

        // Default number of rows to freeze
        int freezeRows = 0;

        // Load and parse the configuration file
        try
        {
            if (File.Exists(configPath))
            {
                string json = File.ReadAllText(configPath);
                using JsonDocument doc = JsonDocument.Parse(json);
                if (doc.RootElement.TryGetProperty("FreezeRows", out JsonElement element) &&
                    element.TryGetInt32(out int value))
                {
                    freezeRows = value;
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error reading configuration: {ex.Message}");
        }

        // Path to the input workbook
        const string inputPath = "input.xlsx";

        // Verify that the input file exists before loading
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0]; // work with the first worksheet

            // Apply freeze panes based on the configuration
            // FreezePanes(row, column, totalRows, totalColumns) freezes rows above 'row' and columns left of 'column'.
            if (freezeRows > 0)
            {
                // 'row' is the first unfrozen row (1‑based). 'totalRows' is the number of rows to freeze.
                // We freeze only rows, so totalColumns is set to 0.
                sheet.FreezePanes(freezeRows + 1, 1, freezeRows, 0);
            }

            // Save the modified workbook
            const string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing workbook: {ex.Message}");
        }
    }
}
