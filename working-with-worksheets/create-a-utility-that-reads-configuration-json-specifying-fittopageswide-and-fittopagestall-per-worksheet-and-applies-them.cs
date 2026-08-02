// Title: Configure per‑worksheet FitToPages settings from JSON using Aspose.Cells for .NET
// Description: A C# console utility that loads an Excel workbook, reads a JSON file describing worksheet names or indexes with desired FitToPagesWide and FitToPagesTall values, disables percent scaling, applies the settings via PageSetup.SetFitToPages, and saves the updated file.
// Keywords: Aspose.Cells FitToPages JSON | C# set worksheet print scaling | PageSetup SetFitToPages example | load workbook apply config | batch worksheet page layout
// Common Searches: set FitToPagesWide per worksheet Aspose.Cells | read JSON to configure Excel page setup C# | apply print scaling to multiple sheets programmatically | Aspose.Cells page setup from configuration file
// Developer Intent: Read a JSON configuration that maps worksheet identifiers to FitToPagesWide/Tall values and programmatically apply those print‑scaling settings to the matching sheets in an Excel workbook.
// Use Cases: Adjust the print layout of generated reports automatically based on a configurable JSON template. | Integrate the utility into a CI/CD pipeline to ensure consistent page scaling across all worksheets before distribution. | Extend the JSON schema to include additional PageSetup options (e.g., PrintArea, Orientation) and apply them in a single pass.
// AI Prompts: Write C# code that parses a JSON file and sets FitToPagesWide and FitToPagesTall for each worksheet in an Aspose.Cells workbook, supporting lookup by name or index. | Show how to add a PrintArea field to the JSON model and update the utility to apply it alongside FitToPages settings. | Suggest robust error‑handling strategies for missing or duplicate worksheet references when applying page‑setup configurations.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace FitToPagesUtility
{
    // Model representing the JSON configuration for a worksheet
    // A C# console utility that loads an Excel workbook, reads a JSON file describing worksheet names or indexes with desired FitToPagesWide and FitToPagesTall values, disables percent scaling, applies the settings via PageSetup.SetFitToPages, and saves the updated file.
    public class WorksheetFitConfig
    {
        public string Name { get; set; }               // Worksheet name (optional, can be null)
        public int? Index { get; set; }                // Worksheet index (optional, can be null)
        public int FitToPagesWide { get; set; }        // Desired pages wide
        public int FitToPagesTall { get; set; }        // Desired pages tall
    }

    // Root object for deserialization
    public class FitConfigRoot
    {
        public List<WorksheetFitConfig> Worksheets { get; set; }
    }

    public static class Program
    {
        // Entry point
        public static void Main(string[] args)
        {
            // Expect two arguments: input workbook path and configuration JSON path
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: FitToPagesUtility <inputWorkbook> <configJson> [outputWorkbook]");
                return;
            }

            string inputWorkbookPath = args[0];
            string configJsonPath = args[1];
            string outputWorkbookPath = args.Length > 2 ? args[2] : "output.xlsx";

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputWorkbookPath);

            // Read and deserialize the JSON configuration
            string jsonContent = File.ReadAllText(configJsonPath);
            FitConfigRoot configRoot = JsonSerializer.Deserialize<FitConfigRoot>(jsonContent);

            if (configRoot?.Worksheets == null)
            {
                Console.WriteLine("Invalid configuration file.");
                return;
            }

            // Apply FitToPages settings per worksheet
            foreach (var wsConfig in configRoot.Worksheets)
            {
                Worksheet worksheet = null;

                // Resolve worksheet by name if provided
                if (!string.IsNullOrEmpty(wsConfig.Name))
                {
                    worksheet = workbook.Worksheets[wsConfig.Name];
                }
                // Otherwise resolve by index if provided
                else if (wsConfig.Index.HasValue && wsConfig.Index.Value >= 0 && wsConfig.Index.Value < workbook.Worksheets.Count)
                {
                    worksheet = workbook.Worksheets[wsConfig.Index.Value];
                }

                if (worksheet == null)
                {
                    Console.WriteLine($"Worksheet not found (Name='{wsConfig.Name}', Index={wsConfig.Index}). Skipping.");
                    continue;
                }

                // Ensure scaling uses FitToPages rather than percent scale
                worksheet.PageSetup.IsPercentScale = false;

                // Apply the fit-to-pages settings (rule: SetFitToPages)
                worksheet.PageSetup.SetFitToPages(wsConfig.FitToPagesWide, wsConfig.FitToPagesTall);
            }

            // Save the modified workbook (lifecycle rule: save)
            workbook.Save(outputWorkbookPath);
            Console.WriteLine($"Workbook saved to '{outputWorkbookPath}'.");
        }
    }
}
