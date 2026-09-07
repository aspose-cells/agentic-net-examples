// Title: Validate CSS linear-gradient generated from Aspose.Cells WordArt fills against original gradient stops using C#
// AI Prompts: Write a C# method that uses reflection to read the GradientStopCollection from an Aspose.Cells FillFormat and returns a list of (Color, double) tuples. | Create a function that builds a CSS linear-gradient string from a list of color‑position pairs, formatting each color as an rgba value with correct alpha handling. | Develop a verification routine that parses a CSS gradient string back into color‑position tuples and compares them with the original list, allowing a small tolerance for position differences.
// Common Searches: how to read WordArt gradient stops from Aspose.Cells using C# reflection | convert Aspose.Cells gradient fill to CSS linear-gradient string in .NET | compare original gradient stops with generated CSS gradient in C# | parse rgba values from a CSS gradient string using C# | validate CSS gradient generated from Excel WordArt shapes
// Tags: Aspose.Cells gradient stop extraction via reflection | C# build CSS linear-gradient from WordArt fill | validate CSS gradient against original Aspose.Cells stops | System.Drawing.Color to CSS rgba conversion | parse CSS linear-gradient string in .NET

using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace WordArtGradientValidation
{
    // The example loads an Excel workbook, iterates through its shapes, extracts gradient stops from WordArt fills via reflection, converts those stops into a CSS linear-gradient string with rgba colors, parses the CSS back into color‑position tuples, and verifies that the parsed stops match the original ones (including alpha) within a tiny tolerance, outputting PASS or FAIL for each shape.
    class WordArtGradientValidator
    {
        // Convert System.Drawing.Color to CSS rgba string.
        private static string ColorToRgba(Color color)
        {
            double alpha = color.A / 255.0;
            return $"rgba({color.R},{color.G},{color.B},{alpha.ToString(CultureInfo.InvariantCulture)})";
        }

        // Generate CSS linear-gradient definition from a list of gradient stops.
        private static string GenerateCssGradient(List<(Color color, double position)> stops)
        {
            const string direction = "to bottom";
            var parts = new List<string>();

            foreach (var (color, position) in stops)
            {
                double positionPercent = position * 100.0;
                string rgba = ColorToRgba(color);
                parts.Add($"{rgba} {positionPercent.ToString(CultureInfo.InvariantCulture)}%");
            }

            string stopsPart = string.Join(", ", parts);
            return $"linear-gradient({direction}, {stopsPart})";
        }

        // Parse CSS gradient string back into a list of (Color, Position) tuples.
        private static List<(Color color, double position)> ParseCssGradient(string css)
        {
            var result = new List<(Color, double)>();

            int start = css.IndexOf('(');
            int end = css.LastIndexOf(')');
            if (start < 0 || end < 0 || end <= start) return result;

            string inner = css.Substring(start + 1, end - start - 1);
            int firstComma = inner.IndexOf(',');
            if (firstComma < 0) return result;

            string stopsPart = inner.Substring(firstComma + 1).Trim();
            string[] stopTokens = stopsPart.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (string token in stopTokens)
            {
                string trimmed = token.Trim();

                int rgbaStart = trimmed.IndexOf("rgba", StringComparison.OrdinalIgnoreCase);
                int rgbaEnd = trimmed.IndexOf(')', rgbaStart);
                if (rgbaStart < 0 || rgbaEnd < 0) continue;

                string rgbaContent = trimmed.Substring(rgbaStart + 5, rgbaEnd - rgbaStart - 5);
                string[] rgbaParts = rgbaContent.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                if (rgbaParts.Length != 4) continue;

                int r = int.Parse(rgbaParts[0].Trim(), CultureInfo.InvariantCulture);
                int g = int.Parse(rgbaParts[1].Trim(), CultureInfo.InvariantCulture);
                int b = int.Parse(rgbaParts[2].Trim(), CultureInfo.InvariantCulture);
                double a = double.Parse(rgbaParts[3].Trim(), CultureInfo.InvariantCulture);
                byte alpha = (byte)Math.Round(a * 255);

                Color color = Color.FromArgb(alpha, r, g, b);

                string afterRgba = trimmed.Substring(rgbaEnd + 1).Trim();
                if (afterRgba.EndsWith("%"))
                    afterRgba = afterRgba.Substring(0, afterRgba.Length - 1);

                double position = double.Parse(afterRgba, CultureInfo.InvariantCulture) / 100.0;
                result.Add((color, position));
            }

            return result;
        }

        // Validate that CSS gradient matches the original gradient stops.
        private static bool ValidateGradient(List<(Color color, double position)> originalStops, string cssGradient)
        {
            var parsedStops = ParseCssGradient(cssGradient);

            if (originalStops.Count != parsedStops.Count)
                return false;

            for (int i = 0; i < originalStops.Count; i++)
            {
                var (origColor, origPos) = originalStops[i];
                var (parsedColor, parsedPos) = parsedStops[i];

                if (origColor.ToArgb() != parsedColor.ToArgb())
                    return false;

                const double tolerance = 0.0001;
                if (Math.Abs(origPos - parsedPos) > tolerance)
                    return false;
            }

            return true;
        }

        // Retrieve gradient stops via reflection to stay compatible with different Aspose.Cells versions.
        private static List<(Color color, double position)> GetGradientStops(FillFormat fill)
        {
            var stops = new List<(Color, double)>();

            try
            {
                var prop = fill.GetType().GetProperty("GradientStopCollection");
                if (prop == null) return stops;

                var collection = prop.GetValue(fill) as IEnumerable;
                if (collection == null) return stops;

                foreach (var stop in collection)
                {
                    var colorProp = stop.GetType().GetProperty("Color");
                    var posProp = stop.GetType().GetProperty("Position");
                    if (colorProp == null || posProp == null) continue;

                    var sysColor = (Color)colorProp.GetValue(stop);
                    var position = Convert.ToDouble(posProp.GetValue(stop));
                    stops.Add((sysColor, position));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to retrieve gradient stops: {ex.Message}");
            }

            return stops;
        }

        static void Main()
        {
            const string inputPath = "WordArtSample.xlsx";

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                Workbook workbook = new Workbook(inputPath);
                Worksheet sheet = workbook.Worksheets[0];

                foreach (Shape shape in sheet.Shapes)
                {
                    try
                    {
                        FillFormat fill = shape.Fill;

                        if (fill == null || fill.FillType != FillType.Gradient)
                        {
                            Console.WriteLine($"Shape '{shape.Name}' does not have a gradient fill.");
                            continue;
                        }

                        var stops = GetGradientStops(fill);
                        if (stops.Count == 0)
                        {
                            Console.WriteLine($"Shape '{shape.Name}' has no gradient stops to process.");
                            continue;
                        }

                        string cssGradient = GenerateCssGradient(stops);
                        Console.WriteLine($"Generated CSS Gradient for shape '{shape.Name}': {cssGradient}");

                        bool isValid = ValidateGradient(stops, cssGradient);
                        Console.WriteLine($"Validation result for shape '{shape.Name}': {(isValid ? "PASS" : "FAIL")}");
                    }
                    catch (Exception shapeEx)
                    {
                        Console.WriteLine($"Error processing shape '{shape.Name}': {shapeEx.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during processing: {ex.Message}");
            }
        }
    }
}
