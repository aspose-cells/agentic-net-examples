// Title: Command‑line utility to set worksheet FitToPagesWide and FitToPagesTall from a JSON file using Aspose.Cells for .NET
// Common Searches: how to set worksheet page fit (FitToPagesWide/FitToPagesTall) from a JSON config with Aspose.Cells C# | Aspose.Cells C# read JSON mapping sheet names to page setup and apply FitToPages settings | C# command line tool to update Excel print scaling per sheet using a JSON file and Aspose.Cells | load workbook and apply per‑sheet FitToPagesWide/FitToPagesTall values from configuration file in .NET | using System.Text.Json to configure Aspose.Cells page setup for multiple worksheets
// Tags: Aspose.Cells worksheet FitToPagesWide configuration | Aspose.Cells worksheet FitToPagesTall configuration | JSON based page setup update for Excel using Aspose.Cells | C# command‑line tool for Excel print scaling per sheet | System.Text.Json deserialization for Aspose.Cells page settings

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace FitToPagesUtility
{
    // Model representing the configuration for a single worksheet
    // A C# console application that accepts an Excel file path and a JSON configuration file, deserializes the JSON into a dictionary of worksheet names with FitToPagesWide and FitToPagesTall values, loads the workbook with Aspose.Cells, applies the specified page‑fit settings to each matching worksheet via PageSetup, and saves the workbook, handling missing sheets and I/O errors.
    public class WorksheetFitConfig
    {
        public int FitToPagesWide { get; set; }
        public int FitToPagesTall { get; set; }
    }

    // Model representing the whole configuration JSON
    // Expected format:
    // {
    //   "Sheet1": { "FitToPagesWide": 1, "FitToPagesTall": 0 },
    //   "Sheet2": { "FitToPagesWide": 2, "FitToPagesTall": 2 }
    // }
    public class FitConfig : Dictionary<string, WorksheetFitConfig>
    {
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Expect two arguments: path to the Excel file and path to the JSON config file
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: FitToPagesUtility <excelFilePath> <configJsonPath>");
                return;
            }

            string excelPath = args[0];
            string jsonPath = args[1];

            // Load configuration JSON
            FitConfig config;
            try
            {
                string jsonContent = File.ReadAllText(jsonPath);
                config = JsonSerializer.Deserialize<FitConfig>(jsonContent, new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                });
                if (config == null)
                {
                    Console.WriteLine("Configuration file is empty or invalid.");
                    return;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to read configuration: {ex.Message}");
                return;
            }

            // Load the workbook using Aspose.Cells
            Workbook workbook;
            try
            {
                workbook = new Workbook(excelPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Apply FitToPages settings per worksheet
            foreach (var kvp in config)
            {
                string sheetName = kvp.Key;
                WorksheetFitConfig sheetConfig = kvp.Value;

                // Find the worksheet by name; if not found, skip
                Worksheet sheet = workbook.Worksheets[sheetName];
                if (sheet == null)
                {
                    Console.WriteLine($"Worksheet \"{sheetName}\" not found in workbook. Skipping.");
                    continue;
                }

                // Apply the FitToPages settings
                sheet.PageSetup.FitToPagesWide = sheetConfig.FitToPagesWide;
                sheet.PageSetup.FitToPagesTall = sheetConfig.FitToPagesTall;
                Console.WriteLine($"Applied FitToPagesWide={sheetConfig.FitToPagesWide}, FitToPagesTall={sheetConfig.FitToPagesTall} to \"{sheetName}\".");
            }

            // Save the modified workbook (overwrite original)
            try
            {
                workbook.Save(excelPath);
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
