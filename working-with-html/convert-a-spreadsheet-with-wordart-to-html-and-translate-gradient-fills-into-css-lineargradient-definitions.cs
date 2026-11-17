// Title: Convert Excel WordArt with gradient fills to HTML and embed CSS linear‑gradient backgrounds using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export WordArt shapes with gradient fill to HTML | How to translate Excel gradient fill into CSS linear‑gradient in C# | Replace Aspose generated <img> tags with background‑gradient style in HTML output | Save workbook as HTML and preserve WordArt gradient colors using Aspose.Cells | Build CSS linear‑gradient string from FillFormat in Aspose.Cells example
// Tags: Aspose.Cells convert WordArt gradient to CSS | C# export Excel shapes to HTML with CSS backgrounds | HTMLSaveOptions gradient fill handling Aspose | Build CSS linear-gradient from FillFormat C# | Inject CSS gradient into Aspose HTML output | WordArt shape processing Aspose.Cells .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an Excel workbook, scans each worksheet for WordArt shapes that use gradient fills, converts those fills into CSS linear‑gradient strings, saves the workbook to HTML in memory, and then injects the generated gradient as a style attribute on the corresponding <img> elements before writing the final HTML file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Store CSS gradient definitions for each WordArt shape
            var shapeGradients = new Dictionary<string, string>();

            // Iterate through all worksheets and their shapes
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Shape shape in sheet.Shapes)
                {
                    // Process only WordArt shapes
                    if (shape.IsWordArt)
                    {
                        FillFormat fill = shape.Fill;

                        // Check if the fill is a gradient
                        if (fill != null && fill.FillType == FillType.Gradient)
                        {
                            try
                            {
                                // Convert the Aspose gradient to a CSS linear‑gradient string
                                string cssGradient = BuildCssLinearGradient(fill);
                                // Use the shape's name as a unique identifier
                                shapeGradients[shape.Name] = cssGradient;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Failed to process gradient for shape '{shape.Name}': {ex.Message}");
                            }
                        }
                    }
                }
            }

            // Save the workbook to HTML in memory
            var htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            using (var ms = new MemoryStream())
            {
                workbook.Save(ms, htmlOptions);
                ms.Position = 0;
                string html = new StreamReader(ms).ReadToEnd();

                // Inject the CSS gradient definitions into the generated HTML
                foreach (var kvp in shapeGradients)
                {
                    // Aspose renders shapes as <img> tags with an id attribute equal to the shape name
                    // Example: <img id="WordArt_0" ...>
                    // Replace the id attribute with a style attribute that contains the gradient
                    string idPattern = $"id=\"{kvp.Key}\"";
                    string replacement = $"id=\"{kvp.Key}\" style=\"background:{kvp.Value};\"";
                    html = html.Replace(idPattern, replacement);
                }

                // Write the final HTML to disk
                File.WriteAllText(outputPath, html);
                Console.WriteLine($"HTML output written to: {outputPath}");
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper method: converts an Aspose gradient fill into a CSS linear‑gradient definition
    private static string BuildCssLinearGradient(FillFormat fill)
    {
        // Determine gradient direction
        string direction = "to bottom";

        // Prefer explicit degree if set
        if (fill.GradientDegree > 0)
        {
            direction = $"{fill.GradientDegree}deg";
        }
        // If no explicit degree, fallback to a generic direction (could be enhanced with more mapping)
        else
        {
            direction = "to bottom";
        }

        // Use the two gradient colors provided by FillFormat
        string color1 = $"#{fill.GradientColor1.R:X2}{fill.GradientColor1.G:X2}{fill.GradientColor1.B:X2}";
        string color2 = $"#{fill.GradientColor2.R:X2}{fill.GradientColor2.G:X2}{fill.GradientColor2.B:X2}";

        return $"linear-gradient({direction}, {color1}, {color2})";
    }
}
