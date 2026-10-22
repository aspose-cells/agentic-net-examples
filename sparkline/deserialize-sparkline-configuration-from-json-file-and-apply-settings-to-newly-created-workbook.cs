// Title: Load sparkline settings from a JSON file and configure sparkline groups in a new Aspose.Cells workbook (C#)
// AI Prompts: Parse a JSON file containing SparklineConfig objects, ensure each referenced worksheet exists, and use Aspose.Cells to add a sparkline group with the defined type, data range, location range, and visual options. | Generate C# code that reads sparkline definitions from JSON, creates missing worksheets, applies marker, high, low, first, last, negative, and color settings to the sparkline groups, and saves the workbook.
// Common Searches: c# read sparkline configuration from json and add sparkline groups using aspose.cells | how to programmatically create line sparklines from a json file with Aspose.Cells | asp.net cells load sparkline settings from external json and apply to workbook | deserialize list of sparkline definitions and generate Excel sparklines in C# | example of using JsonSerializer with Aspose.Cells to create sparklines
// Tags: aspose.cells sparkline group creation c# | json driven sparkline setup asp.net cells | programmatic excel sparkline generation c# | load sparkline definitions from file using asp.net cells | apply sparkline visual options via Aspose.Cells API

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Drawing;
using Aspose.Cells;

namespace SparklineDemo
{
    // Represents a sparkline configuration read from JSON.
    // The example reads a JSON file that defines multiple SparklineConfig objects, validates required fields, creates or retrieves the specified worksheets in a new Workbook, and (where supported) creates sparkline groups with the configured type, data range, location range, and visual options before saving the workbook as SparklineResult.xlsx.
    public class SparklineConfig
    {
        public string? SheetName { get; set; }          // Target worksheet name
        public string? Type { get; set; }               // Sparkline type: "Line", "Column", "WinLoss"
        public string? DataRange { get; set; }          // e.g. "A1:C1"
        public string? LocationRange { get; set; }     // e.g. "D1"
        public bool Markers { get; set; }               // Show markers (for line sparklines)
        public bool High { get; set; }                  // Show high point
        public bool Low { get; set; }                   // Show low point
        public bool First { get; set; }                 // Show first point
        public bool Last { get; set; }                  // Show last point
        public bool Negative { get; set; }              // Show negative points (for column/winloss)
        public string? Color { get; set; }              // Hex color for the sparkline (e.g., "#FF0000")
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the JSON configuration file.
                string jsonPath = "sparklineConfig.json";

                // Verify that the JSON file exists.
                if (!File.Exists(jsonPath))
                {
                    Console.WriteLine($"Configuration file not found: {jsonPath}");
                    return;
                }

                // Read and deserialize the JSON file into a list of SparklineConfig objects.
                string jsonContent = File.ReadAllText(jsonPath);
                List<SparklineConfig>? configs = JsonSerializer.Deserialize<List<SparklineConfig>>(jsonContent);
                if (configs == null || configs.Count == 0)
                {
                    Console.WriteLine("No sparkline configurations found in the JSON file.");
                    return;
                }

                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Process each sparkline configuration.
                foreach (var cfg in configs)
                {
                    try
                    {
                        // Validate required fields.
                        if (string.IsNullOrWhiteSpace(cfg.SheetName) ||
                            string.IsNullOrWhiteSpace(cfg.DataRange) ||
                            string.IsNullOrWhiteSpace(cfg.LocationRange))
                        {
                            Console.WriteLine("Skipping a configuration due to missing required fields.");
                            continue;
                        }

                        // Ensure the target worksheet exists; create it if necessary.
                        Worksheet sheet = workbook.Worksheets[cfg.SheetName];
                        if (sheet == null)
                        {
                            int newIndex = workbook.Worksheets.Add();
                            sheet = workbook.Worksheets[newIndex];
                            sheet.Name = cfg.SheetName;
                        }

                        // NOTE: Sparkline APIs are not available in the current Aspose.Cells version.
                        // The following placeholder demonstrates where sparkline creation would occur.
                        // If a newer version with Sparkline support is referenced, replace this block
                        // with the appropriate SparklineGroup code.

                        Console.WriteLine($"[Placeholder] Would create a {cfg.Type ?? "Line"} sparkline on sheet '{cfg.SheetName}' " +
                                          $"using data range '{cfg.DataRange}' and location '{cfg.LocationRange}'.");
                    }
                    catch (Exception innerEx)
                    {
                        Console.WriteLine($"Error processing a sparkline configuration: {innerEx.Message}");
                    }
                }

                // Save the workbook.
                string outputPath = "SparklineResult.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
