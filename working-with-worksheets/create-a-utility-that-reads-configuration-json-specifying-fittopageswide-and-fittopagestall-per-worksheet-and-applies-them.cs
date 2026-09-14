// Title: C# command‑line utility to apply FitToPagesWide and FitToPagesTall per worksheet from a JSON configuration using Aspose.Cells
// AI Prompts: Generate C# code that reads a JSON file mapping worksheet names to FitToPagesWide and FitToPagesTall values, opens an Excel workbook with Aspose.Cells, updates each sheet's PageSetup, and saves the file. | Create a console application that accepts two arguments (Excel file path and JSON config path) and programmatically sets the print scaling for each worksheet based on the supplied configuration.
// Common Searches: how to programmatically set FitToPagesWide for each sheet in Aspose.Cells C# | apply different FitToPagesTall values to multiple worksheets using a JSON file | C# console app to batch modify Excel page setup with Aspose.Cells | load external configuration to change print scaling of Excel sheets in .NET | Aspose.Cells command line tool for per‑sheet page layout settings
// Tags: Aspose.Cells set worksheet FitToPagesWide | JSON driven page setup update Aspose.Cells | C# batch modify Excel print scaling | command line utility Aspose.Cells page layout | per sheet print scaling automation .NET

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace FitToPagesUtility
{
    // Represents the configuration for a single worksheet.
    // A C# console utility that reads a JSON configuration mapping worksheet names to FitToPagesWide and FitToPagesTall values, loads the specified Excel workbook with Aspose.Cells, applies the defined page‑setup settings to each matching worksheet, and saves the workbook.
    public class WorksheetFitConfig
    {
        public int FitToPagesWide { get; set; }
        public int FitToPagesTall { get; set; }
    }

    // Represents the overall configuration mapping worksheet names to their settings.
    public class FitConfig
    {
        public Dictionary<string, WorksheetFitConfig> Worksheets { get; set; } = new Dictionary<string, WorksheetFitConfig>();
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Expect two arguments: path to the Excel file and path to the JSON config file.
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: FitToPagesUtility <excelFilePath> <configJsonPath>");
                return;
            }

            string excelPath = args[0];
            string configPath = args[1];

            // Load configuration from JSON.
            FitConfig config = LoadConfig(configPath);

            // Load the workbook using Aspose.Cells.
            Workbook workbook = new Workbook(excelPath);

            // Apply FitToPages settings per worksheet as defined in the config.
            foreach (var kvp in config.Worksheets)
            {
                string sheetName = kvp.Key;
                WorksheetFitConfig sheetConfig = kvp.Value;

                // Find the worksheet by name; skip if not found.
                Worksheet sheet = workbook.Worksheets[sheetName];
                if (sheet == null)
                {
                    Console.WriteLine($"Worksheet \"{sheetName}\" not found in the workbook. Skipping.");
                    continue;
                }

                // Apply the FitToPages settings.
                sheet.PageSetup.FitToPagesWide = sheetConfig.FitToPagesWide;
                sheet.PageSetup.FitToPagesTall = sheetConfig.FitToPagesTall;
            }

            // Save the modified workbook (overwrites the original file).
            workbook.Save(excelPath);
            Console.WriteLine("FitToPages settings applied and workbook saved successfully.");
        }

        // Helper method to deserialize the JSON configuration file.
        private static FitConfig LoadConfig(string configPath)
        {
            string json = File.ReadAllText(configPath);
            // The JSON structure should match FitConfig, e.g.:
            // {
            //   "Worksheets": {
            //     "Sheet1": { "FitToPagesWide": 1, "FitToPagesTall": 2 },
            //     "Sheet2": { "FitToPagesWide": 0, "FitToPagesTall": 1 }
            //   }
            // }
            var options = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            };
            return JsonSerializer.Deserialize<FitConfig>(json, options) ?? new FitConfig();
        }
    }
}
