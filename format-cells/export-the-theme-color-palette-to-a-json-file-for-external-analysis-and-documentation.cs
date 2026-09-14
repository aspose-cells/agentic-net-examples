// Title: Export Excel workbook theme color palette to a formatted JSON file with Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx file using Aspose.Cells, iterates over every ThemeColorType, converts each color to a hex string, and saves the collection as an indented JSON array. | Enhance the export routine to include separate numeric R, G, and B fields for each theme color alongside the hex representation in the JSON output. | Add robust error handling that verifies the source workbook exists, creates the target directory if missing, and logs any exceptions that occur during the JSON export.
// Common Searches: how to get Excel theme colors as hex values using Aspose.Cells in C# | export Aspose.Cells workbook theme palette to JSON file for documentation | C# code sample to list ThemeColorType colors from an .xlsx and write to JSON | save Excel theme color collection to formatted JSON with Aspose.Cells .NET
// Tags: Aspose.Cells export theme palette to JSON | C# retrieve Excel theme colors hex | serialize ThemeColorType collection as JSON | write workbook theme colors to formatted JSON file | Aspose.Cells theme color extraction .NET

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// The program loads an Excel workbook, iterates through all ThemeColorType values, converts each theme color to a hex string, and writes the resulting list to a pretty‑printed JSON file, creating the output directory when necessary.
class ExportThemeColors
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "themeColors.json";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve theme colors using the current Aspose.Cells API
            var colors = new List<object>();
            foreach (ThemeColorType themeType in Enum.GetValues(typeof(ThemeColorType)))
            {
                Color clr = workbook.GetThemeColor(themeType);
                string hex = $"#{clr.A:X2}{clr.R:X2}{clr.G:X2}{clr.B:X2}";
                colors.Add(new { Theme = themeType.ToString(), Hex = hex });
            }

            // Serialize the list to JSON with indentation for readability
            var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
            string json = JsonSerializer.Serialize(colors, jsonOptions);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the JSON to a file
            File.WriteAllText(outputPath, json);

            Console.WriteLine($"Theme colors have been exported to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
