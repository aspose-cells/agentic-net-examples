// Title: Export Excel WordArt with gradient fills to HTML and create individual CSS gradient classes using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx workbook with Aspose.Cells, scans every worksheet for WordArt shapes that use a gradient fill, reads the gradient colors (handling both GradientColor and GradientColor1/2 via reflection), generates a distinct CSS class containing a linear‑gradient rule for each shape, saves the workbook as HTML, writes the CSS to an external .css file, and inserts a <link> tag referencing the stylesheet into the HTML. | Implement a routine that maps Shape objects to unique class names, converts System.Drawing.Color values to hex strings, appends `.class { background: linear-gradient(to right, #..., #...); }` statements to a StringBuilder, and integrates the resulting stylesheet with the HTML output produced by Aspose.Cells.
// Common Searches: how to export Excel WordArt with gradient colors to HTML using Aspose.Cells | c# generate separate CSS classes for each WordArt gradient when converting workbook to HTML | using reflection to read GradientColor properties of shapes in Aspose.Cells | add external CSS file to HTML saved by Aspose.Cells workbook | extract gradient fill colors from Excel shapes in .NET
// Tags: Aspose.Cells HTML export with custom CSS for shapes | C# read shape fill properties from Excel workbook | generate unique CSS selectors for Excel shape objects | use .NET metadata inspection to access shape fill data in Aspose.Cells | apply CSS gradient backgrounds to exported Excel content

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads an Excel workbook, iterates through all worksheets and shapes, identifies WordArt objects with gradient fills, retrieves their gradient colors (using reflection for newer and older property names), creates a unique CSS class with a linear‑gradient background for each shape, saves the workbook as HTML, writes the CSS rules to an external file, and injects a <link> tag referencing the stylesheet into the generated HTML.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string htmlPath = "Output.html";
            const string cssPath = "wordart-gradients.css";

            // Verify that the input workbook exists.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: File \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook that contains WordArt objects.
            Workbook workbook = new Workbook(inputPath);

            // Prepare a StringBuilder to collect CSS class definitions for each gradient.
            StringBuilder cssBuilder = new StringBuilder();

            // Map each shape to a unique CSS class name.
            Dictionary<Shape, string> shapeToClass = new Dictionary<Shape, string>();
            int classCounter = 1;

            // Iterate through all worksheets and their shapes.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Shape shape in sheet.Shapes)
                {
                    try
                    {
                        // Process only shapes that have a gradient fill.
                        FillFormat fill = shape.Fill;
                        if (fill == null || fill.FillType != FillType.Gradient)
                            continue;

                        // Create a distinct class name for this shape.
                        string className = $"wordart-grad-{classCounter++}";
                        shapeToClass[shape] = className;

                        Color color1 = Color.Empty;
                        Color color2 = Color.Empty;

                        // Attempt to read newer GradientColor property via reflection.
                        try
                        {
                            var gradientObj = fill.GetType().GetProperty("GradientColor")?.GetValue(fill, null);
                            if (gradientObj != null)
                            {
                                var prop1 = gradientObj.GetType().GetProperty("Color1")?.GetValue(gradientObj, null);
                                var prop2 = gradientObj.GetType().GetProperty("Color2")?.GetValue(gradientObj, null);
                                if (prop1 is Color c1 && prop2 is Color c2)
                                {
                                    color1 = c1;
                                    color2 = c2;
                                }
                            }
                        }
                        catch { /* ignore reflection errors */ }

                        // Fallback to older GradientColor1/GradientColor2 properties.
                        if (color1.IsEmpty || color2.IsEmpty)
                        {
                            try
                            {
                                var prop1 = fill.GetType().GetProperty("GradientColor1")?.GetValue(fill, null);
                                var prop2 = fill.GetType().GetProperty("GradientColor2")?.GetValue(fill, null);
                                if (prop1 is Color c1 && prop2 is Color c2)
                                {
                                    color1 = c1;
                                    color2 = c2;
                                }
                            }
                            catch { /* ignore */ }
                        }

                        // If colors were retrieved, build the CSS rule.
                        if (!color1.IsEmpty && !color2.IsEmpty)
                        {
                            const string direction = "to right"; // Simplified gradient direction.
                            cssBuilder.AppendLine(
                                $".{className} {{ background: linear-gradient({direction}, {ColorToHex(color1)}, {ColorToHex(color2)}); }}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to process shape: {ex.Message}");
                    }
                }
            }

            // Configure HTML save options.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportActiveWorksheetOnly = false // Export the whole workbook.
            };

            // Save the workbook as HTML.
            try
            {
                workbook.Save(htmlPath, htmlOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save HTML: {ex.Message}");
                return;
            }

            // Write the generated CSS to an external file.
            try
            {
                File.WriteAllText(cssPath, cssBuilder.ToString());
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write CSS file: {ex.Message}");
            }

            // Inject a reference to the CSS file into the generated HTML.
            if (File.Exists(htmlPath))
            {
                try
                {
                    string htmlContent = File.ReadAllText(htmlPath);
                    string cssLinkTag = $"<link rel=\"stylesheet\" type=\"text/css\" href=\"{cssPath}\" />";
                    htmlContent = htmlContent.Replace("<head>", $"<head>{Environment.NewLine}{cssLinkTag}");
                    File.WriteAllText(htmlPath, htmlContent);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to inject CSS link: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }

    // Helper method to convert a System.Drawing.Color to a hex string.
    private static string ColorToHex(Color color)
    {
        return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
    }
}
