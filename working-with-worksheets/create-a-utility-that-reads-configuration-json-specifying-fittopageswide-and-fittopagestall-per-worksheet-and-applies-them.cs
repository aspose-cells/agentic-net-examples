// Title: C# CLI utility to set per‑sheet FitToPagesWide/FitToPagesTall from JSON with Aspose.Cells
// Description: A command‑line tool that loads an Excel workbook, reads a JSON file containing worksheet names with FitToPagesWide and FitToPagesTall values, applies PageSetup.SetFitToPages (case‑insensitive), disables percent scaling, and saves the updated file.
// Keywords: Aspose.Cells | C# | FitToPagesWide | FitToPagesTall | JSON configuration | page setup | command line utility | Excel print scaling | batch worksheet settings | GitHub example
// Common Searches: set FitToPagesWide per worksheet Aspose.Cells C# | apply FitToPagesTall from JSON to Excel sheets | C# program to configure page setup using JSON | Aspose.Cells command line tool for print scaling | case insensitive worksheet name matching JSON Aspose
// Developer Intent: Create a reusable CLI program that reads a JSON map of worksheet names to FitToPagesWide/FitToPagesTall values and applies those page‑setup settings to an Excel workbook via Aspose.Cells.
// Use Cases: Automate print‑layout adjustments for dozens of workbooks based on a central JSON template. | Generate printable reports where each sheet requires a distinct page count without hard‑coding values. | Integrate into CI/CD pipelines to enforce consistent page‑setup standards before distribution.
// AI Prompts: Generate code to validate that FitToPagesWide and FitToPagesTall are positive integers before applying them. | Write unit tests that confirm sheets matching the JSON receive the correct settings while others remain unchanged. | Add logging that warns when a worksheet exists in the workbook but has no entry in the JSON configuration.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace FitToPagesUtility
{
    // Represents the configuration for a single worksheet
    // A command‑line tool that loads an Excel workbook, reads a JSON file containing worksheet names with FitToPagesWide and FitToPagesTall values, applies PageSetup.SetFitToPages (case‑insensitive), disables percent scaling, and saves the updated file.
    public class SheetFitConfig
    {
        public string Name { get; set; }          // Worksheet name (case‑insensitive)
        public int FitToPagesWide { get; set; }   // Number of pages wide
        public int FitToPagesTall { get; set; }   // Number of pages tall
    }

    // Root object of the JSON configuration
    public class FitToPagesConfig
    {
        public List<SheetFitConfig> Sheets { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Expect three arguments: input workbook path, config json path, output workbook path
            if (args.Length != 3)
            {
                Console.WriteLine("Usage: FitToPagesUtility <input.xlsx> <config.json> <output.xlsx>");
                return;
            }

            string workbookPath = args[0];
            string configPath   = args[1];
            string outputPath   = args[2];

            // Load the workbook (creation / loading rule)
            Workbook workbook = new Workbook(workbookPath);

            // Read and deserialize the JSON configuration
            string json = File.ReadAllText(configPath);
            FitToPagesConfig config = JsonSerializer.Deserialize<FitToPagesConfig>(json,
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

            if (config?.Sheets == null || config.Sheets.Count == 0)
            {
                Console.WriteLine("No sheet configuration found in the JSON file.");
                return;
            }

            // Apply FitToPages settings per worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Find matching configuration by worksheet name (ignore case)
                SheetFitConfig match = config.Sheets.Find(s =>
                    string.Equals(s.Name, sheet.Name, StringComparison.OrdinalIgnoreCase));

                if (match != null)
                {
                    // Use PageSetup.SetFitToPages method (method rule)
                    sheet.PageSetup.SetFitToPages(match.FitToPagesWide, match.FitToPagesTall);
                    // Ensure scaling is based on FitToPages rather than percent scale
                    sheet.PageSetup.IsPercentScale = false;
                }
            }

            // Save the modified workbook (save rule)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
