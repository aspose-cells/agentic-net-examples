// Title: Load Sparkline Settings from JSON and Create Sparklines in an Aspose.Cells Workbook (C#)
// Description: C# sample that reads a JSON file describing sparkline parameters (data range, location, type, orientation, colors, line weight, preset style, high/low markers), falls back to a default when the file is missing or invalid, creates a new Aspose.Cells workbook, populates sample data, builds SparklineGroup objects with the deserialized settings, adds the sparklines, and saves the workbook as an XLSX file.
// Keywords: Aspose.Cells | C# sparkline JSON | deserialize sparkline configuration | create sparkline group programmatically | sparkline series color hex | sparkline line weight | sparkline preset style | high low point sparkline | Excel sparkline automation | JSON to Aspose.Cells
// Common Searches: How to import sparkline settings from JSON using Aspose.Cells for .NET | Create sparklines in C# with Aspose.Cells from a configuration file | Apply sparkline preset style and color from JSON in an Excel workbook | Default sparkline configuration when JSON file is missing Aspose.Cells | Map hex color to CellsColor for sparkline in Aspose.Cells
// Developer Intent: Read a JSON file that defines sparkline properties, generate the corresponding SparklineGroup objects in a new Aspose.Cells workbook, and save the result.
// Use Cases: Enable business users to modify sparkline appearance by editing a JSON file instead of changing code. | Provide a robust fallback configuration so the workbook always contains at least one sparkline. | Support multiple sparkline groups with different types, orientations, and visual styles in a single sheet. | Integrate external configuration pipelines (e.g., CI/CD, data‑driven dashboards) that supply sparkline definitions in JSON.
// AI Prompts: Write C# code that parses a JSON array of sparkline definitions and creates matching SparklineGroup objects in an Aspose.Cells workbook, handling optional fields like SeriesColor and LineWeight. | Generate a method that validates a sparkline JSON file, returns a default configuration on failure, and applies the settings to a newly created workbook. | Show how to extend SparklineConfig with additional marker options (e.g., ShowFirstPoint, ShowLastPoint) and incorporate them into Aspose.Cells sparkline creation.

using System;
using System.Drawing;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineFromJsonDemo
{
    // Classes that map to the JSON structure
    // C# sample that reads a JSON file describing sparkline parameters (data range, location, type, orientation, colors, line weight, preset style, high/low markers), falls back to a default when the file is missing or invalid, creates a new Aspose.Cells workbook, populates sample data, builds SparklineGroup objects with the deserialized settings, adds the sparklines, and saves the workbook as an XLSX file.
    public class SparklineConfig
    {
        public string DataRange { get; set; }               // e.g. "A1:D1"
        public string LocationCell { get; set; }            // e.g. "E1"
        public int Type { get; set; }                       // 0=Line,1=Column,2=Stacked
        public bool IsVertical { get; set; }                // false = by row, true = by column
        public bool ShowHighPoint { get; set; }
        public bool ShowLowPoint { get; set; }
        public string SeriesColor { get; set; }             // HTML hex, e.g. "#FF6600"
        public double? LineWeight { get; set; }             // optional
        public string PresetStyle { get; set; }             // e.g. "Style5"
    }

    public class WorkbookConfig
    {
        public SparklineConfig[] Sparklines { get; set; }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the JSON configuration file
                string jsonPath = "sparklineConfig.json";

                WorkbookConfig config = null;

                // Load configuration if the file exists
                if (File.Exists(jsonPath))
                {
                    try
                    {
                        string json = File.ReadAllText(jsonPath);
                        config = JsonSerializer.Deserialize<WorkbookConfig>(json);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to deserialize JSON config: {ex.Message}");
                    }
                }

                // Fallback to a default configuration when file is missing or invalid
                if (config?.Sparklines == null || config.Sparklines.Length == 0)
                {
                    config = new WorkbookConfig
                    {
                        Sparklines = new[]
                        {
                            new SparklineConfig
                            {
                                DataRange = "A1:D1",
                                LocationCell = "E1",
                                Type = 0,
                                IsVertical = false,
                                ShowHighPoint = true,
                                ShowLowPoint = true,
                                SeriesColor = "#FF6600",
                                LineWeight = 0.5,
                                PresetStyle = "Style5"
                            }
                        }
                    };
                }

                // ---------- Create a new workbook (lifecycle rule: create) ----------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data referenced by the sparkline
                sheet.Cells["A1"].PutValue(5);
                sheet.Cells["B1"].PutValue(2);
                sheet.Cells["C1"].PutValue(1);
                sheet.Cells["D1"].PutValue(3);
                sheet.Cells["A2"].PutValue(7);
                sheet.Cells["B2"].PutValue(4);
                sheet.Cells["C2"].PutValue(6);
                sheet.Cells["D2"].PutValue(2);

                // ---------- Apply sparkline settings from the configuration ----------
                foreach (var sc in config.Sparklines)
                {
                    // Convert the location cell to a CellArea (single cell)
                    CellArea location = CellArea.CreateCellArea(sc.LocationCell, sc.LocationCell);

                    // Add a sparkline group according to the configuration
                    int groupIdx = sheet.SparklineGroups.Add(
                        (SparklineType)sc.Type,
                        sc.DataRange,
                        sc.IsVertical,
                        location);

                    SparklineGroup group = sheet.SparklineGroups[groupIdx];

                    // Apply optional visual settings
                    if (sc.ShowHighPoint) group.ShowHighPoint = true;
                    if (sc.ShowLowPoint)  group.ShowLowPoint = true;

                    if (!string.IsNullOrEmpty(sc.SeriesColor))
                    {
                        CellsColor seriesClr = workbook.CreateCellsColor();
                        seriesClr.Color = ColorTranslator.FromHtml(sc.SeriesColor);
                        group.SeriesColor = seriesClr;
                    }

                    if (sc.LineWeight.HasValue)
                        group.LineWeight = sc.LineWeight.Value;

                    if (!string.IsNullOrEmpty(sc.PresetStyle) &&
                        Enum.TryParse<SparklinePresetStyleType>(sc.PresetStyle, out var preset))
                    {
                        group.PresetStyle = preset;
                    }

                    // Add the actual sparkline to the group
                    group.Sparklines.Add(sc.DataRange, location.StartRow, location.StartColumn);
                }

                // ---------- Save the workbook (lifecycle rule: save) ----------
                string outputPath = "SparklineFromJsonDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
