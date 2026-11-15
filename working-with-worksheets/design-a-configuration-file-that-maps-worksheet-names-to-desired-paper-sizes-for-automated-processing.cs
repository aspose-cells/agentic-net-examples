// Title: Set worksheet‑specific paper sizes from a JSON configuration using Aspose.Cells for .NET
// Common Searches: c# aspocells set paper size per worksheet using json mapping | how to apply different page setup sizes to Excel sheets with Aspose.Cells | load worksheet‑to‑paper‑size dictionary from json in Aspose.Cells .NET | configure page setup for multiple sheets based on external config file aspocells | read json file and set PageSetup.PaperSize for each worksheet in C#
// Tags: Aspose.Cells set worksheet paper size | JSON driven page setup configuration | C# parse PaperSizeType enum from string | Excel workbook batch page setup automation | per‑sheet page setup using external config .NET

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// The example reads a JSON file that maps worksheet names to Aspose.Cells PaperSizeType values, loads an Excel workbook, iterates through each worksheet, and assigns the specified paper size via the sheet's PageSetup.PaperSize property when a matching entry exists. It then saves the updated workbook, handling missing files and invalid enum values.
class WorksheetPaperSizeMapper
{
    // Path to the JSON configuration file that maps worksheet names to paper size names.
    private const string ConfigPath = "worksheetPaperSizes.json";

    // Path to the source Excel workbook.
    private const string WorkbookPath = "input.xlsx";

    // Path where the processed workbook will be saved.
    private const string OutputPath = "output.xlsx";

    static void Main()
    {
        try
        {
            // Verify configuration file exists.
            if (!File.Exists(ConfigPath))
                throw new FileNotFoundException($"Configuration file not found: {ConfigPath}");

            // Load the worksheet‑to‑paper‑size mapping from the configuration file.
            Dictionary<string, PaperSizeType> paperSizeMap = LoadPaperSizeMapping(ConfigPath);

            // Verify workbook file exists.
            if (!File.Exists(WorkbookPath))
                throw new FileNotFoundException($"Workbook file not found: {WorkbookPath}");

            // Load the workbook from disk.
            Workbook workbook = new Workbook(WorkbookPath);

            // Iterate through all worksheets and apply the configured paper size when a match is found.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (paperSizeMap.TryGetValue(sheet.Name, out PaperSizeType size))
                {
                    // Set the paper size for the current worksheet.
                    sheet.PageSetup.PaperSize = size;
                }
            }

            // Save the modified workbook.
            workbook.Save(OutputPath);
            Console.WriteLine($"Workbook saved successfully to '{OutputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Reads the JSON file and converts it to a dictionary of worksheet name -> PaperSizeType enum.
    private static Dictionary<string, PaperSizeType> LoadPaperSizeMapping(string path)
    {
        string json = File.ReadAllText(path);

        // The JSON is expected to be a simple object: { "Sheet1": "A4", "Sheet2": "Letter", ... }
        Dictionary<string, string> rawMap = JsonSerializer.Deserialize<Dictionary<string, string>>(json);

        var result = new Dictionary<string, PaperSizeType>(StringComparer.OrdinalIgnoreCase);
        if (rawMap == null)
            return result;

        foreach (KeyValuePair<string, string> kvp in rawMap)
        {
            // Try to parse the string value to the PaperSizeType enum (case‑insensitive).
            if (Enum.TryParse<PaperSizeType>(kvp.Value, true, out PaperSizeType size))
            {
                result[kvp.Key] = size;
            }
            else
            {
                // Invalid paper size strings are ignored, but a warning is emitted.
                Console.WriteLine($"Warning: Invalid paper size '{kvp.Value}' for worksheet '{kvp.Key}'.");
            }
        }

        return result;
    }
}
