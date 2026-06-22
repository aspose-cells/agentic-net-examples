using System;
using System.IO;
using System.Text.Json;
using System.Collections.Generic;
using Aspose.Cells;

namespace FitToPagesUtility
{
    // Represents a single worksheet configuration from the JSON file
    public class WorksheetFitConfig
    {
        // Optional zero‑based worksheet index
        public int? Index { get; set; }

        // Optional worksheet name
        public string Name { get; set; }

        // Desired number of pages wide
        public int FitToPagesWide { get; set; }

        // Desired number of pages tall
        public int FitToPagesTall { get; set; }
    }

    public class Program
    {
        // args[0] = input workbook path
        // args[1] = JSON configuration path
        // args[2] = output workbook path
        public static void Main(string[] args)
        {
            if (args.Length < 3)
            {
                Console.WriteLine("Usage: FitToPagesUtility <input.xlsx> <config.json> <output.xlsx>");
                return;
            }

            string workbookPath = args[0];
            string jsonPath = args[1];
            string outputPath = args[2];

            // Load the workbook (create/load rule)
            Workbook workbook = new Workbook(workbookPath);

            // Read and deserialize the JSON configuration
            string json = File.ReadAllText(jsonPath);
            List<WorksheetFitConfig> configs = JsonSerializer.Deserialize<List<WorksheetFitConfig>>(json);

            // Apply FitToPages settings to each specified worksheet
            foreach (var cfg in configs)
            {
                Worksheet ws = null;

                // Prefer index if supplied and valid
                if (cfg.Index.HasValue && cfg.Index.Value >= 0 && cfg.Index.Value < workbook.Worksheets.Count)
                {
                    ws = workbook.Worksheets[cfg.Index.Value];
                }
                // Fallback to name lookup
                else if (!string.IsNullOrEmpty(cfg.Name))
                {
                    ws = workbook.Worksheets[cfg.Name];
                }

                if (ws == null)
                {
                    Console.WriteLine($"Worksheet not found (Index={cfg.Index}, Name={cfg.Name}). Skipping.");
                    continue;
                }

                // Apply the FitToPages settings using the documented method
                ws.PageSetup.SetFitToPages(cfg.FitToPagesWide, cfg.FitToPagesTall);
                // Ensure the worksheet uses FitToPages scaling rather than percent scaling
                ws.PageSetup.IsPercentScale = false;
            }

            // Save the modified workbook (save rule)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}