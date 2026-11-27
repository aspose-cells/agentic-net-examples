// Title: Load Sparkline configuration from a JSON file and apply it to a SparklineGroup in an Aspose.Cells workbook (C#)
// Common Searches: C# read sparkline settings from JSON and create SparklineGroup with Aspose.Cells | how to map JSON sparkline properties to Aspose.Cells SparklineGroup in .NET | populate missing data range for sparkline when using Aspose.Cells workbook | convert cell address string to CellArea for sparkline location in Aspose.Cells C# | set sparkline colors from HTML hex values using Aspose.Cells API
// Tags: json sparkline configuration deserialization Aspose.Cells | create sparklinegroup with data range Aspose.Cells | apply sparkline color properties from HTML Aspose.Cells | cell address to CellArea conversion Aspose.Cells | auto-fill sparkline data range cells Aspose.Cells

using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineFromJson
{
    // Define classes that match the expected JSON structure.
    // Only the properties that are needed for SparklineGroup configuration are included.
    // The example reads a JSON file containing Sparkline configuration, deserializes it into a SparklineConfig object, creates a new Workbook, optionally fills missing cells in the specified data range, parses the sparkline type and location address, adds a SparklineGroup to the worksheet, applies boolean, color, line weight, and plot settings from the configuration, and saves the workbook as an XLSX file.
    public class SparklineConfig
    {
        [JsonPropertyName("type")]
        public string Type { get; set; }

        [JsonPropertyName("dataRange")]
        public string DataRange { get; set; }

        [JsonPropertyName("location")]
        public string Location { get; set; }

        [JsonPropertyName("showFirstPoint")]
        public bool? ShowFirstPoint { get; set; }

        [JsonPropertyName("showHighPoint")]
        public bool? ShowHighPoint { get; set; }

        [JsonPropertyName("showLowPoint")]
        public bool? ShowLowPoint { get; set; }

        [JsonPropertyName("showNegativePoints")]
        public bool? ShowNegativePoints { get; set; }

        [JsonPropertyName("showMarkers")]
        public bool? ShowMarkers { get; set; }

        [JsonPropertyName("showLastPoint")]
        public bool? ShowLastPoint { get; set; }

        [JsonPropertyName("firstPointColor")]
        public string FirstPointColor { get; set; }

        [JsonPropertyName("highPointColor")]
        public string HighPointColor { get; set; }

        [JsonPropertyName("lowPointColor")]
        public string LowPointColor { get; set; }

        [JsonPropertyName("negativePointsColor")]
        public string NegativePointsColor { get; set; }

        [JsonPropertyName("markersColor")]
        public string MarkersColor { get; set; }

        [JsonPropertyName("lastPointColor")]
        public string LastPointColor { get; set; }

        [JsonPropertyName("seriesColor")]
        public string SeriesColor { get; set; }

        [JsonPropertyName("lineWeight")]
        public double? LineWeight { get; set; }

        [JsonPropertyName("plotEmptyCellsType")]
        public string PlotEmptyCellsType { get; set; }

        [JsonPropertyName("plotRightToLeft")]
        public bool? PlotRightToLeft { get; set; }

        [JsonPropertyName("displayHidden")]
        public bool? DisplayHidden { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Determine JSON file path.
            string jsonPath = args.Length > 0 ? args[0] : "sparklineConfig.json";

            if (!File.Exists(jsonPath))
            {
                Console.Error.WriteLine($"Error: JSON configuration file not found at '{jsonPath}'.");
                return;
            }

            SparklineConfig config;
            try
            {
                string jsonContent = File.ReadAllText(jsonPath);
                var options = new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true,
                    ReadCommentHandling = JsonCommentHandling.Skip,
                    AllowTrailingCommas = true
                };
                config = JsonSerializer.Deserialize<SparklineConfig>(jsonContent, options);
                if (config == null)
                {
                    Console.Error.WriteLine("Error: Deserialized configuration is null.");
                    return;
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error reading or parsing JSON file: {ex.Message}");
                return;
            }

            // Create a new workbook.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data if the referenced range does not exist.
            try
            {
                PopulateSampleDataIfNeeded(sheet, config.DataRange);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error while populating sample data: {ex.Message}");
                return;
            }

            // Convert string type to SparklineType enum.
            SparklineType sparklineType = ParseSparklineType(config.Type);

            // Determine the location cell area for the sparkline group.
            CellArea locationArea = ParseLocationToCellArea(config.Location, sheet);

            // Add the sparkline group.
            int groupIndex;
            try
            {
                // The Add method signature:
                // int Add(SparklineType type, string dataRange, bool isDateAxis, CellArea location)
                groupIndex = sheet.SparklineGroups.Add(sparklineType, config.DataRange, false, locationArea);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error adding sparkline group: {ex.Message}");
                return;
            }

            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Apply configuration values to the SparklineGroup.
            ApplyGroupSettings(group, config, workbook);

            // Save the workbook.
            string outputPath = "output_sparkline.xlsx";
            try
            {
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }

        private static SparklineType ParseSparklineType(string typeStr)
        {
            if (Enum.TryParse<SparklineType>(typeStr, true, out var result))
                return result;

            Console.Error.WriteLine($"Warning: Unable to parse SparklineType '{typeStr}'. Defaulting to Line.");
            return SparklineType.Line;
        }

        private static CellArea ParseLocationToCellArea(string location, Worksheet sheet)
        {
            if (string.IsNullOrWhiteSpace(location))
            {
                location = "E1";
            }

            try
            {
                int row = CellReferenceHelper.CellNameToRow(location);
                int col = CellReferenceHelper.CellNameToColumn(location);
                return new CellArea
                {
                    StartRow = row,
                    EndRow = row,
                    StartColumn = col,
                    EndColumn = col
                };
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error parsing location '{location}': {ex.Message}. Using default E1.");
                int row = CellReferenceHelper.CellNameToRow("E1");
                int col = CellReferenceHelper.CellNameToColumn("E1");
                return new CellArea { StartRow = row, EndRow = row, StartColumn = col, EndColumn = col };
            }
        }

        private static void ApplyGroupSettings(SparklineGroup group, SparklineConfig cfg, Workbook wb)
        {
            // Helper to convert a color string to CellsColor.
            CellsColor GetCellsColor(string colorStr)
            {
                if (string.IsNullOrWhiteSpace(colorStr))
                    return null;

                try
                {
                    Color sysColor = ColorTranslator.FromHtml(colorStr);
                    CellsColor cellsColor = wb.CreateCellsColor();
                    cellsColor.Color = sysColor;
                    return cellsColor;
                }
                catch
                {
                    try
                    {
                        Color sysColor = Color.FromName(colorStr);
                        if (!sysColor.IsKnownColor && !sysColor.IsNamedColor)
                            throw new Exception();
                        CellsColor cellsColor = wb.CreateCellsColor();
                        cellsColor.Color = sysColor;
                        return cellsColor;
                    }
                    catch
                    {
                        Console.Error.WriteLine($"Warning: Unable to parse color '{colorStr}'. Ignoring.");
                        return null;
                    }
                }
            }

            // Boolean properties.
            if (cfg.ShowFirstPoint.HasValue) group.ShowFirstPoint = cfg.ShowFirstPoint.Value;
            if (cfg.ShowHighPoint.HasValue) group.ShowHighPoint = cfg.ShowHighPoint.Value;
            if (cfg.ShowLowPoint.HasValue) group.ShowLowPoint = cfg.ShowLowPoint.Value;
            if (cfg.ShowNegativePoints.HasValue) group.ShowNegativePoints = cfg.ShowNegativePoints.Value;
            if (cfg.ShowMarkers.HasValue) group.ShowMarkers = cfg.ShowMarkers.Value;
            if (cfg.ShowLastPoint.HasValue) group.ShowLastPoint = cfg.ShowLastPoint.Value;
            if (cfg.DisplayHidden.HasValue) group.DisplayHidden = cfg.DisplayHidden.Value;
            if (cfg.PlotRightToLeft.HasValue) group.PlotRightToLeft = cfg.PlotRightToLeft.Value;
            if (cfg.LineWeight.HasValue) group.LineWeight = cfg.LineWeight.Value;

            // Color properties.
            var firstPtClr = GetCellsColor(cfg.FirstPointColor);
            if (firstPtClr != null) group.FirstPointColor = firstPtClr;

            var highPtClr = GetCellsColor(cfg.HighPointColor);
            if (highPtClr != null) group.HighPointColor = highPtClr;

            var lowPtClr = GetCellsColor(cfg.LowPointColor);
            if (lowPtClr != null) group.LowPointColor = lowPtClr;

            var negPtClr = GetCellsColor(cfg.NegativePointsColor);
            if (negPtClr != null) group.NegativePointsColor = negPtClr;

            var markersClr = GetCellsColor(cfg.MarkersColor);
            if (markersClr != null) group.MarkersColor = markersClr;

            var lastPtClr = GetCellsColor(cfg.LastPointColor);
            if (lastPtClr != null) group.LastPointColor = lastPtClr;

            var seriesClr = GetCellsColor(cfg.SeriesColor);
            if (seriesClr != null) group.SeriesColor = seriesClr;

            // PlotEmptyCellsType handling.
            if (!string.IsNullOrWhiteSpace(cfg.PlotEmptyCellsType))
            {
                if (Enum.TryParse<PlotEmptyCellsType>(cfg.PlotEmptyCellsType, true, out var emptyType))
                {
                    group.PlotEmptyCellsType = emptyType;
                }
                else
                {
                    Console.Error.WriteLine($"Warning: Invalid PlotEmptyCellsType '{cfg.PlotEmptyCellsType}'. Ignored.");
                }
            }
        }

        private static void PopulateSampleDataIfNeeded(Worksheet sheet, string dataRange)
        {
            if (string.IsNullOrWhiteSpace(dataRange))
                return;

            // Split the range into start and end addresses.
            string[] parts = dataRange.Split(':');
            if (parts.Length != 2)
                return; // Invalid range format; nothing to populate.

            CellArea area = CellArea.CreateCellArea(parts[0], parts[1]);
            Cells cells = sheet.Cells;
            int counter = 1;
            for (int row = area.StartRow; row <= area.EndRow; row++)
            {
                for (int col = area.StartColumn; col <= area.EndColumn; col++)
                {
                    if (cells[row, col].Value == null)
                    {
                        cells[row, col].PutValue(counter++);
                    }
                }
            }
        }
    }

    // Helper class to translate cell addresses to row/column indices.
    internal static class CellReferenceHelper
    {
        public static int CellNameToRow(string cellName)
        {
            string rowPart = "";
            foreach (char ch in cellName)
            {
                if (char.IsDigit(ch))
                    rowPart += ch;
            }
            if (int.TryParse(rowPart, out int rowNumber))
                return rowNumber - 1; // zero‑based index
            throw new FormatException($"Cannot parse row from cell name '{cellName}'.");
        }

        public static int CellNameToColumn(string cellName)
        {
            string colPart = "";
            foreach (char ch in cellName)
            {
                if (char.IsLetter(ch))
                    colPart += ch;
                else
                    break;
            }
            colPart = colPart.ToUpperInvariant();
            int sum = 0;
            foreach (char ch in colPart)
            {
                sum *= 26;
                sum += (ch - 'A' + 1);
            }
            return sum - 1; // zero‑based index
        }
    }
}
