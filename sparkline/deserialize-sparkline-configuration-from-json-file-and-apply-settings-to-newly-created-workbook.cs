using System;
using System.Drawing;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineConfigDemo
{
    // Classes that represent the JSON structure for sparkline configuration
    public class SparklineConfig
    {
        public string Type { get; set; }                     // "Line", "Column", or "Stacked"
        public string DataRange { get; set; }                // e.g., "A1:D1"
        public bool IsVertical { get; set; }                 // true = plot by column, false = by row
        public CellAreaConfig LocationRange { get; set; }    // where the sparkline will be placed
        public string SeriesColor { get; set; }              // hex color, e.g., "#FF8000"
        public bool ShowHighPoint { get; set; }
        public string HighPointColor { get; set; }
        public bool ShowLowPoint { get; set; }
        public string LowPointColor { get; set; }
        public bool ShowFirstPoint { get; set; }
        public string FirstPointColor { get; set; }
        public bool ShowLastPoint { get; set; }
        public string LastPointColor { get; set; }
        public bool ShowMarkers { get; set; }
        public string MarkersColor { get; set; }
        public bool ShowNegativePoints { get; set; }
        public string NegativePointsColor { get; set; }
        public bool DisplayHidden { get; set; }
        public double LineWeight { get; set; }
        public string PresetStyle { get; set; }              // e.g., "Style5"
        public string PlotEmptyCellsType { get; set; }       // e.g., "Zero"
    }

    public class CellAreaConfig
    {
        public int StartRow { get; set; }
        public int EndRow { get; set; }
        public int StartColumn { get; set; }
        public int EndColumn { get; set; }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the JSON file that contains sparkline configuration
                string configPath = "sparklineConfig.json";

                // Verify that the configuration file exists
                if (!File.Exists(configPath))
                {
                    Console.WriteLine($"Configuration file not found: {Path.GetFullPath(configPath)}");
                    return;
                }

                // Deserialize the JSON configuration
                SparklineConfig config = JsonSerializer.Deserialize<SparklineConfig>(File.ReadAllText(configPath));
                if (config == null)
                {
                    Console.WriteLine("Failed to deserialize configuration.");
                    return;
                }

                // ---------- Create a new workbook (create rule) ----------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some sample data that the sparkline will reference
                sheet.Cells["A1"].PutValue(5);
                sheet.Cells["B1"].PutValue(2);
                sheet.Cells["C1"].PutValue(1);
                sheet.Cells["D1"].PutValue(3);
                sheet.Cells["E1"].PutValue(4);

                // Convert the location range from the config into a CellArea object
                CellArea location = new CellArea
                {
                    StartRow = config.LocationRange.StartRow,
                    EndRow = config.LocationRange.EndRow,
                    StartColumn = config.LocationRange.StartColumn,
                    EndColumn = config.LocationRange.EndColumn
                };

                // Parse the sparkline type enum from string
                SparklineType sparklineType = (SparklineType)Enum.Parse(typeof(SparklineType), config.Type, true);

                // Add a sparkline group using the loaded configuration (add method)
                int groupIdx = sheet.SparklineGroups.Add(sparklineType, config.DataRange, config.IsVertical, location);
                SparklineGroup group = sheet.SparklineGroups[groupIdx];

                // Apply visual settings from the configuration
                if (!string.IsNullOrEmpty(config.SeriesColor))
                {
                    CellsColor seriesClr = workbook.CreateCellsColor();
                    seriesClr.Color = ColorTranslator.FromHtml(config.SeriesColor);
                    group.SeriesColor = seriesClr;
                }

                group.ShowHighPoint = config.ShowHighPoint;
                if (config.ShowHighPoint && !string.IsNullOrEmpty(config.HighPointColor))
                {
                    CellsColor highClr = workbook.CreateCellsColor();
                    highClr.Color = ColorTranslator.FromHtml(config.HighPointColor);
                    group.HighPointColor = highClr;
                }

                group.ShowLowPoint = config.ShowLowPoint;
                if (config.ShowLowPoint && !string.IsNullOrEmpty(config.LowPointColor))
                {
                    CellsColor lowClr = workbook.CreateCellsColor();
                    lowClr.Color = ColorTranslator.FromHtml(config.LowPointColor);
                    group.LowPointColor = lowClr;
                }

                group.ShowFirstPoint = config.ShowFirstPoint;
                if (config.ShowFirstPoint && !string.IsNullOrEmpty(config.FirstPointColor))
                {
                    CellsColor firstClr = workbook.CreateCellsColor();
                    firstClr.Color = ColorTranslator.FromHtml(config.FirstPointColor);
                    group.FirstPointColor = firstClr;
                }

                group.ShowLastPoint = config.ShowLastPoint;
                if (config.ShowLastPoint && !string.IsNullOrEmpty(config.LastPointColor))
                {
                    CellsColor lastClr = workbook.CreateCellsColor();
                    lastClr.Color = ColorTranslator.FromHtml(config.LastPointColor);
                    group.LastPointColor = lastClr;
                }

                group.ShowMarkers = config.ShowMarkers;
                if (config.ShowMarkers && !string.IsNullOrEmpty(config.MarkersColor))
                {
                    CellsColor markersClr = workbook.CreateCellsColor();
                    markersClr.Color = ColorTranslator.FromHtml(config.MarkersColor);
                    group.MarkersColor = markersClr;
                }

                group.ShowNegativePoints = config.ShowNegativePoints;
                if (config.ShowNegativePoints && !string.IsNullOrEmpty(config.NegativePointsColor))
                {
                    CellsColor negClr = workbook.CreateCellsColor();
                    negClr.Color = ColorTranslator.FromHtml(config.NegativePointsColor);
                    group.NegativePointsColor = negClr;
                }

                group.DisplayHidden = config.DisplayHidden;
                group.LineWeight = config.LineWeight;

                if (!string.IsNullOrEmpty(config.PresetStyle))
                {
                    group.PresetStyle = (SparklinePresetStyleType)Enum.Parse(typeof(SparklinePresetStyleType), config.PresetStyle, true);
                }

                if (!string.IsNullOrEmpty(config.PlotEmptyCellsType))
                {
                    group.PlotEmptyCellsType = (PlotEmptyCellsType)Enum.Parse(typeof(PlotEmptyCellsType), config.PlotEmptyCellsType, true);
                }

                // ---------- Save the workbook (save rule) ----------
                string outputPath = "SparklineConfigured.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}