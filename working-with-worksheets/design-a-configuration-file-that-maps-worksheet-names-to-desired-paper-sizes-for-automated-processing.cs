// Title: Set Worksheet‑Specific Paper Sizes from JSON with Aspose.Cells for .NET
// Description: Loads a JSON file that maps worksheet names to Aspose.Cells PaperSizeType values, opens an Excel workbook, applies the defined page‑setup paper size to each matching sheet, and saves the updated file. Includes validation for missing files and invalid enum entries.
// Keywords: Aspose.Cells C# set paper size | worksheet page setup Aspose | JSON configuration Excel paper size | PaperSizeType enum Aspose.Cells | load workbook apply page setup | C# Excel automation paper size | Aspose.Cells per‑sheet printing settings
// Common Searches: how to assign different paper sizes to Excel sheets using Aspose.Cells | apply JSON‑based page setup to worksheets in C# | map worksheet names to PaperSizeType enum Aspose | set page setup paper size programmatically Aspose.Cells | C# read JSON and change Excel sheet print settings
// Developer Intent: Configure individual worksheet print dimensions from an external JSON file.
// Use Cases: Standardize printing formats across multiple reports by defining sheet‑to‑paper‑size mappings in a JSON file. | Automate batch processing of workbooks where each worksheet requires a specific page size (e.g., A4 for data sheets, Letter for summaries). | Gracefully handle missing configuration or incorrect enum names, logging warnings without breaking the workflow.
// AI Prompts: Generate C# code that reads a JSON file of worksheet‑paper‑size pairs and sets PageSetup.PaperSize for each matching sheet using Aspose.Cells. | Create a sample paperSizeConfig.json containing common sizes (A4, Letter, Legal) and explain how Enum.TryParse converts the strings to PaperSizeType. | Suggest robust error‑handling patterns for absent configuration files, unreadable JSON, and invalid PaperSizeType values when working with Aspose.Cells.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// Loads a JSON file that maps worksheet names to Aspose.Cells PaperSizeType values, opens an Excel workbook, applies the defined page‑setup paper size to each matching sheet, and saves the updated file. Includes validation for missing files and invalid enum entries.
class WorksheetPaperSizeConfigurator
{
    static void Main()
    {
        try
        {
            // Path to the JSON configuration file that maps worksheet names to paper size names.
            string configPath = "paperSizeConfig.json";

            // Path to the source workbook that will be processed.
            string inputWorkbookPath = "input.xlsx";

            // Path where the modified workbook will be saved.
            string outputWorkbookPath = "output.xlsx";

            // Load the mapping from the configuration file (if it exists).
            Dictionary<string, PaperSizeType> sheetPaperSizeMap = new Dictionary<string, PaperSizeType>(StringComparer.OrdinalIgnoreCase);
            if (File.Exists(configPath))
            {
                sheetPaperSizeMap = LoadPaperSizeMapping(configPath);
            }
            else
            {
                Console.WriteLine($"Configuration file not found: {configPath}. No paper size changes will be applied.");
            }

            // Verify the input workbook exists before loading.
            if (!File.Exists(inputWorkbookPath))
            {
                Console.WriteLine($"Input workbook not found: {inputWorkbookPath}");
                return;
            }

            // Load the workbook (lifecycle rule: load).
            Workbook workbook = new Workbook(inputWorkbookPath);

            // Iterate through each worksheet and apply the configured paper size if present.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (sheetPaperSizeMap.TryGetValue(sheet.Name, out PaperSizeType paperSize))
                {
                    // Set the paper size for the current worksheet (feature rule: PageSetup.PaperSize).
                    sheet.PageSetup.PaperSize = paperSize;
                    Console.WriteLine($"Worksheet '{sheet.Name}' paper size set to {paperSize}.");
                }
            }

            // Save the modified workbook (lifecycle rule: save).
            workbook.Save(outputWorkbookPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outputWorkbookPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Reads a JSON file where keys are worksheet names and values are paper size enum names.
    // Example JSON:
    // {
    //   "Sheet1": "PaperA4",
    //   "Report": "PaperLetter"
    // }
    static Dictionary<string, PaperSizeType> LoadPaperSizeMapping(string filePath)
    {
        string jsonContent = File.ReadAllText(filePath);
        var rawMapping = JsonSerializer.Deserialize<Dictionary<string, string>>(jsonContent);

        var result = new Dictionary<string, PaperSizeType>(StringComparer.OrdinalIgnoreCase);
        foreach (var kvp in rawMapping)
        {
            // Convert the string value to the corresponding PaperSizeType enum.
            if (Enum.TryParse<PaperSizeType>(kvp.Value, ignoreCase: true, out PaperSizeType size))
            {
                result[kvp.Key] = size;
            }
            else
            {
                Console.WriteLine($"Warning: Invalid paper size '{kvp.Value}' for worksheet '{kvp.Key}'. Entry ignored.");
            }
        }
        return result;
    }
}
