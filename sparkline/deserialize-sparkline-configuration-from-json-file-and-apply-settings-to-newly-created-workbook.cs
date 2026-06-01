using System;
using System.IO;
using System.Text.Json;
using System.Drawing;
using System.Globalization;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineFromJsonDemo
{
    // Represents the JSON structure for a sparkline group configuration
    public class SparklineConfig
    {
        public string? DataRange { get; set; }               // e.g. "A1:D1"
        public string? LocationStart { get; set; }           // e.g. "E1"
        public string? LocationEnd { get; set; }             // e.g. "E1" (same cell for single sparkline)
        public bool IsVertical { get; set; }                 // false = by row, true = by column
        public string? Type { get; set; }                    // "Line", "Column", "Stacked"
        public string? SeriesColor { get; set; }             // Hex color, e.g. "#FF6600"
        public bool ShowHighPoint { get; set; }
        public string? HighPointColor { get; set; }          // Hex color
        public bool ShowLowPoint { get; set; }
        public string? LowPointColor { get; set; }           // Hex color
        public bool ShowFirstPoint { get; set; }
        public string? FirstPointColor { get; set; }         // Hex color
        public bool ShowLastPoint { get; set; }
        public string? LastPointColor { get; set; }          // Hex color
        public bool ShowMarkers { get; set; }
        public string? MarkersColor { get; set; }            // Hex color
        public double LineWeight { get; set; }               // Points
        public string? PresetStyle { get; set; }             // e.g. "Style5"
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the JSON configuration file
                const string jsonPath = "sparklineConfig.json";

                // Ensure the JSON file exists
                if (!File.Exists(jsonPath))
                {
                    Console.WriteLine($"Configuration file not found: {jsonPath}");
                    return;
                }

                // Deserialize the JSON into a SparklineConfig object
                SparklineConfig? config = JsonSerializer.Deserialize<SparklineConfig>(File.ReadAllText(jsonPath));
                if (config == null)
                {
                    Console.WriteLine("Failed to deserialize configuration.");
                    return;
                }

                // ---------- Create a new workbook (lifecycle rule: create) ----------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate the data range referenced in the config (demo data)
                if (!string.IsNullOrEmpty(config.DataRange))
                {
                    PopulateSampleData(sheet, config.DataRange);
                }

                // Convert the string type to the SparklineType enum
                SparklineType sparklineType = config.Type switch
                {
                    "Line" => SparklineType.Line,
                    "Column" => SparklineType.Column,
                    "Stacked" => SparklineType.Stacked,
                    _ => SparklineType.Line
                };

                // Build the location CellArea from start/end cells
                if (string.IsNullOrEmpty(config.LocationStart) || string.IsNullOrEmpty(config.LocationEnd))
                {
                    Console.WriteLine("LocationStart or LocationEnd is missing in configuration.");
                    return;
                }

                CellArea locationArea = CellArea.CreateCellArea(config.LocationStart, config.LocationEnd);

                // Add the sparkline group (lifecycle rule: load – we are loading settings into the workbook)
                int groupIdx = sheet.SparklineGroups.Add(sparklineType, config.DataRange, config.IsVertical, locationArea);
                SparklineGroup group = sheet.SparklineGroups[groupIdx];

                // Apply visual settings from the configuration
                if (!string.IsNullOrEmpty(config.SeriesColor))
                    group.SeriesColor = CreateCellsColor(workbook, config.SeriesColor);

                group.ShowHighPoint = config.ShowHighPoint;
                if (config.ShowHighPoint && !string.IsNullOrEmpty(config.HighPointColor))
                    group.HighPointColor = CreateCellsColor(workbook, config.HighPointColor);

                group.ShowLowPoint = config.ShowLowPoint;
                if (config.ShowLowPoint && !string.IsNullOrEmpty(config.LowPointColor))
                    group.LowPointColor = CreateCellsColor(workbook, config.LowPointColor);

                group.ShowFirstPoint = config.ShowFirstPoint;
                if (config.ShowFirstPoint && !string.IsNullOrEmpty(config.FirstPointColor))
                    group.FirstPointColor = CreateCellsColor(workbook, config.FirstPointColor);

                group.ShowLastPoint = config.ShowLastPoint;
                if (config.ShowLastPoint && !string.IsNullOrEmpty(config.LastPointColor))
                    group.LastPointColor = CreateCellsColor(workbook, config.LastPointColor);

                group.ShowMarkers = config.ShowMarkers;
                if (config.ShowMarkers && !string.IsNullOrEmpty(config.MarkersColor))
                    group.MarkersColor = CreateCellsColor(workbook, config.MarkersColor);

                if (config.LineWeight > 0)
                    group.LineWeight = config.LineWeight;

                if (!string.IsNullOrEmpty(config.PresetStyle) &&
                    Enum.TryParse<SparklinePresetStyleType>(config.PresetStyle, out var preset))
                {
                    group.PresetStyle = preset;
                }

                // ---------- Save the workbook (lifecycle rule: save) ----------
                const string outputPath = "SparklineFromJsonDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Helper to create a CellsColor from a hex string (e.g., "#FF6600")
        private static CellsColor CreateCellsColor(Workbook wb, string hex)
        {
            // Remove leading '#' if present
            if (hex.StartsWith("#"))
                hex = hex.Substring(1);

            // Parse ARGB (8 chars) or RGB (6 chars)
            int argb = hex.Length == 8
                ? int.Parse(hex, NumberStyles.HexNumber)
                : (int.Parse(hex, NumberStyles.HexNumber) | unchecked((int)0xFF000000));

            Color color = Color.FromArgb(argb);
            CellsColor cellsColor = wb.CreateCellsColor();
            cellsColor.Color = color;
            return cellsColor;
        }

        // Simple method to fill the data range with incremental numbers for demo purposes
        private static void PopulateSampleData(Worksheet sheet, string dataRange)
        {
            // Split the range string (e.g., "A1:D1") into start and end cell names
            string[] parts = dataRange.Split(':', StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length != 2)
                throw new ArgumentException($"Invalid data range format: {dataRange}");

            CellArea area = CellArea.CreateCellArea(parts[0], parts[1]);

            int value = 1;
            for (int row = area.StartRow; row <= area.EndRow; row++)
            {
                for (int col = area.StartColumn; col <= area.EndColumn; col++)
                {
                    sheet.Cells[row, col].PutValue(value++);
                }
            }
        }
    }
}