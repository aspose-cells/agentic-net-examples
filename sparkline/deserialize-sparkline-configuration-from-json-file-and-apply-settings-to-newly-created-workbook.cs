// Title: Read Sparkline Configuration from JSON and Apply It to a New Workbook with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a JSON file into a SparklineConfig object, creates a Workbook, adds a SparklineGroup using the parsed type, data range, orientation, and location, and sets visual properties like series color, high/low point colors, line weight, and preset style. | Extend the sample to handle an array of SparklineConfig entries in the JSON file and add each entry as a separate SparklineGroup on the same worksheet. | Add validation that checks enum values for SparklineType and SparklinePresetStyleType, providing fallback defaults when the JSON contains invalid names.
// Common Searches: how to load sparkline settings from a json file using Aspose.Cells in C# | aspnet create sparkline group from external configuration json | set sparkline series color and preset style programmatically with Aspose.Cells | deserialize sparkline configuration json to Aspose.Cells sparkline group example | c# Aspose.Cells sparkline vertical orientation from json
// Tags: json driven sparkline configuration Aspose.Cells | set sparkline visual properties programmatically C# | parse sparkline type enum from string Aspose.Cells | apply series color using color name Aspose.Cells | load sparkline preset style enum from json

using System;
using System.Drawing;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example reads a SparklineConfig JSON file (or falls back to defaults), creates a new Workbook, adds a sparkline group with the specified type, data range, orientation and location, applies visual settings such as series color, high‑point/low‑point colors, line weight and preset style, and saves the workbook as an .xlsx file.
public class SparklineConfig
{
    public string Type { get; set; }                     // e.g., "Line"
    public string DataRange { get; set; }                // e.g., "A1:D1"
    public bool IsVertical { get; set; }                 // false = by row
    public string LocationStartCell { get; set; }        // e.g., "E1"
    public string LocationEndCell { get; set; }          // e.g., "E1"
    public string SeriesColor { get; set; }              // e.g., "Orange"
    public bool ShowHighPoint { get; set; }
    public string HighPointColor { get; set; }
    public bool ShowLowPoint { get; set; }
    public string LowPointColor { get; set; }
    public double LineWeight { get; set; }
    public string PresetStyle { get; set; }              // e.g., "Style5"
}

public class SparklineFromJsonDemo
{
    private const string JsonFilePath = "sparklineConfig.json";
    private const string OutputFilePath = "SparklineFromJsonDemo.xlsx";

    public static void Run()
    {
        try
        {
            // Load configuration (fallback to defaults if file missing)
            SparklineConfig config = LoadConfig(JsonFilePath);

            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (ensure DataRange has values)
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(1);
            sheet.Cells["D1"].PutValue(3);

            // Convert string values to Aspose.Cells enums/types
            SparklineType sparklineType = (SparklineType)Enum.Parse(typeof(SparklineType), config.Type);
            CellArea location = CellArea.CreateCellArea(config.LocationStartCell, config.LocationEndCell);

            // Add sparkline group
            int groupIndex = sheet.SparklineGroups.Add(sparklineType, config.DataRange, config.IsVertical, location);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Apply visual settings
            if (!string.IsNullOrEmpty(config.SeriesColor))
            {
                CellsColor seriesColor = workbook.CreateCellsColor();
                seriesColor.Color = Color.FromName(config.SeriesColor);
                group.SeriesColor = seriesColor;
            }

            group.ShowHighPoint = config.ShowHighPoint;
            if (config.ShowHighPoint && !string.IsNullOrEmpty(config.HighPointColor))
            {
                CellsColor highColor = workbook.CreateCellsColor();
                highColor.Color = Color.FromName(config.HighPointColor);
                group.HighPointColor = highColor;
            }

            group.ShowLowPoint = config.ShowLowPoint;
            if (config.ShowLowPoint && !string.IsNullOrEmpty(config.LowPointColor))
            {
                CellsColor lowColor = workbook.CreateCellsColor();
                lowColor.Color = Color.FromName(config.LowPointColor);
                group.LowPointColor = lowColor;
            }

            group.LineWeight = config.LineWeight;

            if (!string.IsNullOrEmpty(config.PresetStyle))
            {
                SparklinePresetStyleType style = (SparklinePresetStyleType)Enum.Parse(
                    typeof(SparklinePresetStyleType), config.PresetStyle);
                group.PresetStyle = style;
            }

            // Save the workbook
            workbook.Save(OutputFilePath);
            Console.WriteLine($"Workbook saved successfully to '{OutputFilePath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    private static SparklineConfig LoadConfig(string path)
    {
        if (!File.Exists(path))
        {
            // Return a default configuration when the JSON file is absent
            return new SparklineConfig
            {
                Type = "Line",
                DataRange = "A1:D1",
                IsVertical = false,
                LocationStartCell = "E1",
                LocationEndCell = "E1",
                SeriesColor = "Orange",
                ShowHighPoint = true,
                HighPointColor = "Red",
                ShowLowPoint = true,
                LowPointColor = "Blue",
                LineWeight = 0.75,
                PresetStyle = "Style5"
            };
        }

        string json = File.ReadAllText(path);
        return JsonSerializer.Deserialize<SparklineConfig>(json);
    }
}

class Program
{
    static void Main()
    {
        SparklineFromJsonDemo.Run();
    }
}
