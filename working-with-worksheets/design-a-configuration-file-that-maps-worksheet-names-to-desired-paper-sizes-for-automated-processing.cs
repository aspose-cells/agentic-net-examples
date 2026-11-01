// Title: Use a JSON file to map worksheet names to specific paper sizes with Aspose.Cells for .NET
// Common Searches: how to assign different paper sizes to Excel worksheets using Aspose.Cells C# | load worksheet paper size settings from a JSON configuration in a .NET application | apply per‑sheet PageSetup.PaperSize from a dictionary with Aspose.Cells | serialize and deserialize worksheet paper size mapping to JSON with Aspose.Cells | set A4 and Letter paper sizes for specific sheets programmatically in C#
// Tags: json worksheet paper size mapping Aspose.Cells | set page setup paper size per worksheet c# | deserialize worksheet configuration Aspose.Cells | apply paper size dictionary to workbook | configure workbook sheet paper size via json

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

// Define a configuration class that maps worksheet names to paper sizes
// The example defines a WorksheetPaperSizeConfig class that stores a dictionary of worksheet names and their desired PaperSizeType values, serializes this mapping to a formatted JSON file, reads an existing workbook, deserializes the configuration, iterates through the mappings to set each worksheet's PageSetup.PaperSize, and saves the updated workbook.
public class WorksheetPaperSizeConfig
{
    // The key is the worksheet name, the value is the desired paper size
    public Dictionary<string, PaperSizeType> Mappings { get; set; } = new Dictionary<string, PaperSizeType>();
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // Step 1: Create a sample configuration and persist it to a JSON file
        // -----------------------------------------------------------------
        var config = new WorksheetPaperSizeConfig();
        config.Mappings["Sheet1"] = PaperSizeType.PaperA4;          // Map "Sheet1" to A4
        config.Mappings["Report"] = PaperSizeType.PaperLetter;    // Map "Report" to Letter

        string configPath = "WorksheetPaperSizeConfig.json";

        // Serialize the configuration with indentation for readability
        string json = JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(configPath, json);

        // ---------------------------------------------------------------
        // Step 2: Load an existing workbook that will be processed
        // ---------------------------------------------------------------
        var workbook = new Workbook("Input.xlsx"); // Replace with your actual file path

        // ---------------------------------------------------------------
        // Step 3: Load the configuration from the JSON file
        // ---------------------------------------------------------------
        var loadedConfig = JsonSerializer.Deserialize<WorksheetPaperSizeConfig>(File.ReadAllText(configPath));

        // ---------------------------------------------------------------
        // Step 4: Apply the mapped paper sizes to the corresponding worksheets
        // ---------------------------------------------------------------
        foreach (var mapping in loadedConfig.Mappings)
        {
            // Retrieve the worksheet by name; if it doesn't exist, skip it
            var sheet = workbook.Worksheets[mapping.Key];
            if (sheet != null)
            {
                sheet.PageSetup.PaperSize = mapping.Value;
            }
        }

        // ---------------------------------------------------------------
        // Step 5: Save the modified workbook
        // ---------------------------------------------------------------
        workbook.Save("Output.xlsx"); // Replace with your desired output path
    }
}
