// Title: C# utility to apply worksheet‑specific FitToPagesWide and FitToPagesTall values from a JSON file using Aspose.Cells
// Common Searches: how to programmatically set FitToPagesWide and FitToPagesTall for each sheet with Aspose.Cells in C# | read Excel page‑setup configuration from JSON and apply it using Aspose.Cells .NET | apply per‑worksheet page fit settings from a JSON configuration file in a .NET application | Aspose.Cells example for updating worksheet PageSetup based on external JSON data | C# code to load JSON list of worksheets and set FitToPagesWide/FitToPagesTall in an Excel workbook
// Tags: Aspose.Cells set worksheet FitToPagesWide | Aspose.Cells set worksheet FitToPagesTall | C# deserialize JSON to configure Excel page setup | apply per‑sheet page fit settings with Aspose.Cells | Excel workbook page layout automation using Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsFitToPagesUtility
{
    // Model representing configuration for a worksheet
    // The sample loads an Excel workbook with Aspose.Cells, deserializes a JSON file that maps worksheet names to FitToPagesWide and FitToPagesTall values, updates each matching worksheet's PageSetup accordingly, and saves the modified workbook while handling missing files or worksheets.
    public class WorksheetFitConfig
    {
        public string WorksheetName { get; set; }
        public int FitToPagesWide { get; set; }
        public int FitToPagesTall { get; set; }
    }

    public static class FitToPagesApplier
    {
        /// <param name="inputWorkbookPath">Path to the source Excel file.</param>
        /// <param name="configJsonPath">Path to the JSON configuration file.</param>
        /// <param name="outputWorkbookPath">Path where the modified Excel file will be saved.</param>
        public static void ApplyFitToPages(string inputWorkbookPath, string configJsonPath, string outputWorkbookPath)
        {
            try
            {
                // Verify input files exist
                if (!File.Exists(inputWorkbookPath))
                {
                    Console.WriteLine($"Input workbook not found: {inputWorkbookPath}");
                    return;
                }

                if (!File.Exists(configJsonPath))
                {
                    Console.WriteLine($"Configuration file not found: {configJsonPath}");
                    return;
                }

                // Load the workbook (using Aspose.Cells)
                Workbook workbook = new Workbook(inputWorkbookPath);

                // Read and deserialize the JSON configuration
                string jsonContent = File.ReadAllText(configJsonPath);
                List<WorksheetFitConfig> configs = JsonSerializer.Deserialize<List<WorksheetFitConfig>>(jsonContent);

                if (configs == null)
                {
                    Console.WriteLine("No configuration data found in the JSON file.");
                    return;
                }

                // Apply settings per worksheet
                foreach (WorksheetFitConfig cfg in configs)
                {
                    // Find the worksheet by name; if not found, skip to next config
                    Worksheet sheet = workbook.Worksheets[cfg.WorksheetName];
                    if (sheet == null)
                    {
                        Console.WriteLine($"Worksheet \"{cfg.WorksheetName}\" not found. Skipping.");
                        continue;
                    }

                    // Apply FitToPagesWide and FitToPagesTall
                    sheet.PageSetup.FitToPagesWide = cfg.FitToPagesWide;
                    sheet.PageSetup.FitToPagesTall = cfg.FitToPagesTall;
                }

                // Save the modified workbook
                workbook.Save(outputWorkbookPath);
                Console.WriteLine($"Workbook saved successfully to: {outputWorkbookPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while applying FitToPages settings: {ex.Message}");
            }
        }
    }

    // Example usage
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Paths can be adjusted as needed
                string inputPath = "input.xlsx";
                string configPath = "fitConfig.json";
                string outputPath = "output.xlsx";

                FitToPagesApplier.ApplyFitToPages(inputPath, configPath, outputPath);

                Console.WriteLine("FitToPages settings applied successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
