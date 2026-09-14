// Title: Freeze a configurable number of rows in an Excel worksheet using Aspose.Cells and appsettings.json (C#)
// AI Prompts: Read the "FreezeRows" integer from a JSON configuration file and use Worksheet.FreezePanes to freeze that many rows in the first sheet with Aspose.Cells. | Extend the sample to also read a "FreezeColumns" setting and apply Worksheet.FreezePanes to freeze both rows and columns based on the configuration values. | Add fallback logic that uses a default freeze count when the JSON setting is missing, malformed, or outside an acceptable range.
// Common Searches: c# aspnet read FreezeRows from appsettings.json and apply freeze panes using Aspose.Cells | how to set freeze panes dynamically in Aspose.Cells based on a configuration file | using Worksheet.FreezePanes with values from JSON in a .NET console application | Aspose.Cells example to freeze first N rows read from appsettings.json
// Tags: read FreezeRows from JSON configuration Aspose.Cells | apply Worksheet.FreezePanes dynamically .NET | configure Excel freeze panes at runtime | load or create workbook Aspose.Cells C# | handle missing appsettings.json Aspose.Cells example

using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// The C# program reads a "FreezeRows" integer from an appsettings.json file, loads an existing workbook or creates a new one, and freezes the specified number of rows on the first worksheet using Aspose.Cells' Worksheet.FreezePanes method. It includes error handling for missing files, invalid configuration values, and saves the result to output.xlsx.
class Program
{
    static void Main()
    {
        // Read FreezeRows value from appsettings.json if present
        int rowsToFreeze = 0;
        const string configPath = "appsettings.json";
        try
        {
            if (File.Exists(configPath))
            {
                string json = File.ReadAllText(configPath);
                using JsonDocument doc = JsonDocument.Parse(json);
                if (doc.RootElement.TryGetProperty("FreezeRows", out JsonElement elem) && elem.TryGetInt32(out int value))
                {
                    rowsToFreeze = value;
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to read configuration: {ex.Message}");
        }

        // Load existing workbook or create a new one
        string inputPath = "input.xlsx";
        Workbook workbook;
        try
        {
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Apply freeze panes on the first worksheet
        try
        {
            Worksheet sheet = workbook.Worksheets[0];
            // Freeze the specified number of rows (no columns)
            sheet.FreezePanes(rowsToFreeze + 1, 0, rowsToFreeze, 0);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error applying freeze panes: {ex.Message}");
        }

        // Save the modified workbook
        string outputPath = "output.xlsx";
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
