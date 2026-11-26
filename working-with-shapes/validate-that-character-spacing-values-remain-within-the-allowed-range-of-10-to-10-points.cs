// Title: Validate character spacing (TextOptions.Spacing) in Excel text boxes, chart titles, axis titles, and data labels using Aspose.Cells for .NET
// Common Searches: how to check character spacing limits in Aspose.Cells text boxes C# | validate chart title spacing range Aspose.Cells .NET | C# script to ensure Excel text box spacing is between -10 and 10 points | Aspose.Cells verify text options spacing for data labels in charts | detect invalid character spacing values in workbook using Aspose.Cells | range validation for TextOptions.Spacing property in Aspose.Cells
// Tags: character spacing range enforcement Aspose.Cells | Excel text box character spacing check C# | chart title spacing verification Aspose.Cells | axis title spacing validation .NET | data label spacing audit Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace CharacterSpacingValidator
{
    // The C# utility loads an Excel workbook with Aspose.Cells, iterates through worksheets, and checks that the TextOptions.Spacing value of text boxes, chart titles, axis titles, and data label text stays within the allowed -10 to 10 point range. Violations are collected and reported, and the workbook can be saved after validation.
    class Program
    {
        // Allowed character spacing range (in points)
        private const double MinSpacing = -10.0;
        private const double MaxSpacing = 10.0;

        static void Main(string[] args)
        {
            // ------------------------------------------------------------
            // Input handling
            // ------------------------------------------------------------
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: CharacterSpacingValidator <input.xlsx> [output.xlsx]");
                return;
            }

            string inputPath = args[0];
            string outputPath = args.Length > 1 ? args[1] : "Validated_" + Path.GetFileName(inputPath);

            // ------------------------------------------------------------
            // Verify input file exists
            // ------------------------------------------------------------
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // ------------------------------------------------------------
            // Load workbook
            // ------------------------------------------------------------
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
                Console.WriteLine($"Loaded workbook: {inputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook \"{inputPath}\": {ex.Message}");
                return;
            }

            // ------------------------------------------------------------
            // Validate character spacing in all relevant objects
            // ------------------------------------------------------------
            List<string> violations = new List<string>();

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    ValidateWorksheetTextBoxes(sheet, violations);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error validating text boxes in sheet \"{sheet.Name}\": {ex.Message}");
                }

                try
                {
                    ValidateChartObjects(sheet, violations);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error validating charts in sheet \"{sheet.Name}\": {ex.Message}");
                }
            }

            // ------------------------------------------------------------
            // Report results
            // ------------------------------------------------------------
            if (violations.Count == 0)
            {
                Console.WriteLine("All character spacing values are within the allowed range.");
            }
            else
            {
                Console.WriteLine("Character spacing violations detected:");
                foreach (string v in violations)
                {
                    Console.WriteLine(v);
                }
            }

            // ------------------------------------------------------------
            // Save (optional)
            // ------------------------------------------------------------
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook \"{outputPath}\": {ex.Message}");
            }
        }

        private static void ValidateWorksheetTextBoxes(Worksheet sheet, List<string> violations)
        {
            TextBoxCollection textBoxes = sheet.TextBoxes;
            for (int i = 0; i < textBoxes.Count; i++)
            {
                TextBox tb = textBoxes[i];
                // Use TextOptions.Spacing for character spacing
                double spacing = tb.TextOptions?.Spacing ?? 0.0;

                if (!IsWithinRange(spacing))
                {
                    string location = $"Worksheet \"{sheet.Name}\" TextBox[{i}] (Cell: {tb.UpperLeftRow},{tb.UpperLeftColumn})";
                    violations.Add($"{location}: Spacing = {spacing} (allowed: {MinSpacing}..{MaxSpacing})");
                }
            }
        }

        private static void ValidateChartObjects(Worksheet sheet, List<string> violations)
        {
            foreach (Chart chart in sheet.Charts)
            {
                int chartIdx = sheet.Charts.IndexOf(chart);

                // Chart Title
                if (chart.Title?.TextOptions != null)
                {
                    ValidateTextOptions(chart.Title.TextOptions,
                        $"Worksheet \"{sheet.Name}\" Chart[{chartIdx}] Title", violations);
                }

                // Category (X) Axis Title
                if (chart.CategoryAxis?.Title?.TextOptions != null)
                {
                    ValidateTextOptions(chart.CategoryAxis.Title.TextOptions,
                        $"Worksheet \"{sheet.Name}\" Chart[{chartIdx}] Category Axis Title", violations);
                }

                // Value (Y) Axis Title
                if (chart.ValueAxis?.Title?.TextOptions != null)
                {
                    ValidateTextOptions(chart.ValueAxis.Title.TextOptions,
                        $"Worksheet \"{sheet.Name}\" Chart[{chartIdx}] Value Axis Title", violations);
                }

                // Data Labels (if any)
                for (int s = 0; s < chart.NSeries.Count; s++)
                {
                    Series series = chart.NSeries[s];
                    if (series.DataLabels?.TextOptions != null)
                    {
                        ValidateTextOptions(series.DataLabels.TextOptions,
                            $"Worksheet \"{sheet.Name}\" Chart[{chartIdx}] Series[{s}] DataLabels", violations);
                    }
                }
            }
        }

        private static void ValidateTextOptions(TextOptions options, string context, List<string> violations)
        {
            // TextOptions inherits from Font, use Spacing property
            double spacing = options.Spacing;

            if (!IsWithinRange(spacing))
            {
                violations.Add($"{context}: Spacing = {spacing} (allowed: {MinSpacing}..{MaxSpacing})");
            }
        }

        private static bool IsWithinRange(double value)
        {
            return value >= MinSpacing && value <= MaxSpacing;
        }
    }
}
