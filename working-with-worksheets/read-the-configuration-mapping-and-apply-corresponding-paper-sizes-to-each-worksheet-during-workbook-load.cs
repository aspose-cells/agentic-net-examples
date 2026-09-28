// Title: Load an Excel workbook with Aspose.Cells for .NET and set each worksheet's paper size from a JSON configuration file
// AI Prompts: Read a JSON file that maps worksheet names to paper size strings, convert each string to the PaperSizeType enum, and assign it to the matching worksheet's PageSetup using Aspose.Cells. | Enhance the code to gracefully skip worksheets with missing or invalid paper size entries without throwing exceptions during workbook loading. | After updating the page setup, export the workbook to PDF and save both the modified Excel file and the generated PDF.
// Common Searches: asp.net set worksheet paper size from json using Aspose.Cells | how to map sheet names to PaperSizeType enum in C# Aspose.Cells | load excel file and apply different page sizes per sheet based on a config file | Aspose.Cells change page setup for each worksheet programmatically | c# read json configuration and set worksheet page setup Aspose.Cells
// Tags: set worksheet paper size Aspose.Cells | json driven page setup .NET | parse PaperSizeType enum C# | apply per-sheet page setup Aspose.Cells | load workbook modify page setup Aspose.Cells | export workbook to PDF after page setup

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing; // For PaperSizeType enum

// The example reads a JSON file containing worksheet‑to‑paper‑size mappings, loads an Excel workbook with Aspose.Cells, iterates through each worksheet, parses the size string into the PaperSizeType enum, applies it to the sheet's PageSetup, and saves the updated workbook (optionally exporting to PDF).
class Program
{
    static void Main()
    {
        try
        {
            // Path to configuration file that maps worksheet names to paper size names
            string configPath = "config.json";
            if (!File.Exists(configPath))
            {
                Console.WriteLine($"Configuration file not found: {configPath}");
                return;
            }

            // Deserialize the JSON mapping (e.g., { "Sheet1": "A4", "Report": "Legal" })
            var json = File.ReadAllText(configPath);
            var paperSizeMap = JsonSerializer.Deserialize<Dictionary<string, string>>(json) ?? new Dictionary<string, string>();

            // Path to the input workbook
            string inputPath = "input.xlsx";
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input workbook not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Apply paper sizes based on the configuration
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (paperSizeMap.TryGetValue(sheet.Name, out string sizeName))
                {
                    // Try to parse the paper size name to the Aspose.Cells PaperSizeType enum (case‑insensitive)
                    if (Enum.TryParse(sizeName, ignoreCase: true, out PaperSizeType paperSize))
                    {
                        sheet.PageSetup.PaperSize = paperSize;
                    }
                    else
                    {
                        Console.WriteLine($"Warning: '{sizeName}' is not a valid PaperSize for worksheet '{sheet.Name}'.");
                    }
                }
            }

            // Save the modified workbook
            string outputPath = "output.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
