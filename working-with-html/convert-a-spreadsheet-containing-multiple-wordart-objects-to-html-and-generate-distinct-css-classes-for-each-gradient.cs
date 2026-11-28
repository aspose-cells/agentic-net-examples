// Title: Export Excel WordArt shapes to HTML with individual gradient CSS classes using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export WordArt shapes to HTML with separate CSS gradient classes | How to generate unique CSS classes for each WordArt gradient when converting Excel to HTML | Convert multiple WordArt objects in an .xlsx file to HTML preserving gradient fills | Aspose.Cells C# create linear-gradient text style for WordArt in HTML output | Export Excel WordArt to HTML with inline CSS using Aspose.Cells library
// Tags: Aspose.Cells export WordArt to HTML | C# generate CSS gradient for Excel shape | WordArt gradient CSS class generation | Excel shape to HTML conversion Aspose | linear-gradient text fill Aspose.Cells

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Net;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace WordArtToHtmlDemo
{
    // The example loads an .xlsx workbook with Aspose.Cells, iterates over the first worksheet’s shapes, selects only WordArt objects, extracts their text and gradient fill (or falls back to solid black), creates a unique CSS class for each shape using a linear‑gradient background that matches the WordArt’s angle and colors, writes all CSS definitions into a <style> block, and outputs an HTML file where each WordArt is rendered as a <span> with the generated class, preserving the visual appearance.
    class Program
    {
        static void Main(string[] args)
        {
            // -----------------------------------------------------------------
            // 1. Validate input arguments
            // -----------------------------------------------------------------
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: WordArtToHtmlDemo <input.xlsx> <output.html>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                return;
            }

            try
            {
                // -----------------------------------------------------------------
                // 2. Load the workbook
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook(inputPath);
                Worksheet worksheet = workbook.Worksheets[0];
                ShapeCollection shapes = worksheet.Shapes;

                // -----------------------------------------------------------------
                // 3. Find all WordArt shapes and collect gradient information
                // -----------------------------------------------------------------
                var wordArtInfos = new List<WordArtInfo>();
                int classCounter = 1;

                foreach (Shape shape in shapes)
                {
                    // Only process WordArt objects
                    if (!shape.IsWordArt)
                        continue;

                    // Retrieve the displayed text (WordArt text)
                    string text = shape.Text ?? string.Empty;

                    // Retrieve fill information
                    FillFormat fill = shape.Fill;
                    if (fill == null || fill.FillType != FillType.Gradient)
                    {
                        // Solid fill fallback – use black as default color
                        Color solid = Color.Black;
                        string cssClass = $"wordart-grad-{classCounter++}";
                        wordArtInfos.Add(new WordArtInfo
                        {
                            Text = text,
                            CssClassName = cssClass,
                            CssDefinition = GenerateSolidColorCss(cssClass, solid)
                        });
                        continue;
                    }

                    // Gradient colors (Color is a struct, never null)
                    Color startColor = fill.GradientColor1;
                    Color endColor = fill.GradientColor2;

                    // Gradient angle (degrees). Aspose.Cells uses 0‑360 where 0 is left‑to‑right.
                    double angle = fill.GradientDegree;

                    // Build a unique CSS class name
                    string className = $"wordart-grad-{classCounter++}";

                    // Generate the CSS definition for this gradient
                    string css = GenerateGradientCss(className, startColor, endColor, angle);

                    wordArtInfos.Add(new WordArtInfo
                    {
                        Text = text,
                        CssClassName = className,
                        CssDefinition = css
                    });
                }

                // -----------------------------------------------------------------
                // 4. Build the final HTML document
                // -----------------------------------------------------------------
                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                using (StreamWriter writer = new StreamWriter(outputPath, false, Encoding.UTF8))
                {
                    writer.WriteLine("<!DOCTYPE html>");
                    writer.WriteLine("<html lang=\"en\">");
                    writer.WriteLine("<head>");
                    writer.WriteLine("    <meta charset=\"UTF-8\">");
                    writer.WriteLine("    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">");
                    writer.WriteLine("    <title>WordArt Export</title>");
                    writer.WriteLine("    <style>");
                    writer.WriteLine("        body { font-family: Arial, Helvetica, sans-serif; padding: 20px; }");
                    writer.WriteLine("        .wordart { display: inline-block; font-weight: bold; }");
                    // Write each generated CSS class
                    foreach (var info in wordArtInfos)
                    {
                        writer.WriteLine(info.CssDefinition);
                    }
                    writer.WriteLine("    </style>");
                    writer.WriteLine("</head>");
                    writer.WriteLine("<body>");
                    writer.WriteLine("    <h2>Exported WordArt Objects</h2>");
                    writer.WriteLine("    <div>");

                    foreach (var info in wordArtInfos)
                    {
                        // Escape HTML special characters in the text
                        string escapedText = WebUtility.HtmlEncode(info.Text);
                        writer.WriteLine($"        <span class=\"wordart {info.CssClassName}\">{escapedText}</span><br/>");
                    }

                    writer.WriteLine("    </div>");
                    writer.WriteLine("</body>");
                    writer.WriteLine("</html>");
                }

                Console.WriteLine($"HTML file successfully created at \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during processing:");
                Console.WriteLine(ex.Message);
            }
        }

        // Holds information about a single WordArt shape needed for HTML generation.
        private class WordArtInfo
        {
            public string Text { get; set; } = string.Empty;
            public string CssClassName { get; set; } = string.Empty;
            public string CssDefinition { get; set; } = string.Empty;
        }

        // Generates a CSS rule that applies a linear gradient as a text fill.
        private static string GenerateGradientCss(string className, Color start, Color end, double angle)
        {
            string startHex = ColorToHex(start);
            string endHex = ColorToHex(end);
            string gradient = $"linear-gradient({angle.ToString(CultureInfo.InvariantCulture)}deg, {startHex}, {endHex})";

            return $@"
        .{className} {{
            background: {gradient};
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
            color: transparent;
        }}";
        }

        // Generates a CSS rule for a solid fill (fallback when no gradient is present).
        private static string GenerateSolidColorCss(string className, Color color)
        {
            string hex = ColorToHex(color);
            return $@"
        .{className} {{
            color: {hex};
        }}";
        }

        // Helper to convert a System.Drawing.Color to a CSS hex string (#RRGGBB).
        private static string ColorToHex(Color color)
        {
            return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
        }
    }
}
