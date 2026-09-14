// Title: Extract WordArt Gradient Fill Information from Excel and Render It on HTML Canvas Using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, finds all WordArt shapes that use gradient fills, and collects each shape’s width, height, rotation angle, and gradient stop colors with positions. | Create a JavaScript routine that receives the extracted shape data and draws a matching linear gradient on a <canvas> element, then embed this routine into the HTML produced by Aspose.Cells. | Extend the C# program to also handle radial gradient fills from WordArt shapes and output the corresponding canvas drawing commands.
// Common Searches: how to preserve Excel WordArt gradient fills when converting to HTML with Aspose.Cells | c# extract gradient stop data from WordArt shapes using Aspose.Cells | render Excel WordArt gradients on HTML canvas after saving as HTML | aspnet generate canvas fallback for WordArt gradient fills in exported HTML
// Tags: Aspose.Cells extract WordArt gradient fills | C# export Excel to HTML with canvas fallback | HTML canvas linear gradient from shape properties | gradient stop collection using Fill property | render WordArt gradients on canvas | radial gradient support in Aspose.Cells HTML export

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an Excel workbook with Aspose.Cells, identifies WordArt shapes that use gradient fills, extracts each shape’s dimensions, angle, and gradient stop data, saves the workbook as HTML, and injects <canvas> elements plus JavaScript that redraws the gradients on page load, providing a visual fallback for browsers that cannot render the original WordArt.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Collect information about WordArt shapes that use gradient fills
            var wordArtInfos = new List<WordArtInfo>();

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Shape shape in sheet.Shapes)
                {
                    // Identify WordArt shapes
                    if (shape.IsWordArt)
                    {
                        // Use the modern Fill property
                        var fill = shape.Fill;
                        if (fill != null && fill.FillType == FillType.Gradient && fill.GradientFill != null)
                        {
                            var gradientStops = new List<GradientStopInfo>();
                            foreach (var stopObj in fill.GradientFill.GradientStops)
                            {
                                // Use dynamic to access Color property (covers API variations)
                                dynamic stop = stopObj;
                                Color color = stop.Color;
                                double position = stop.Position;

                                gradientStops.Add(new GradientStopInfo
                                {
                                    Color = color,
                                    Position = position
                                });
                            }

                            wordArtInfos.Add(new WordArtInfo
                            {
                                Id = $"wordart_{wordArtInfos.Count}",
                                Width = shape.Width,
                                Height = shape.Height,
                                Angle = fill.GradientFill.Angle,
                                GradientStops = gradientStops
                            });
                        }
                    }
                }
            }

            // Save workbook as HTML (in-memory)
            var htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            string html;
            using (var ms = new MemoryStream())
            {
                workbook.Save(ms, htmlOptions);
                ms.Position = 0;
                using var reader = new StreamReader(ms);
                html = reader.ReadToEnd();
            }

            // Build canvas elements for each WordArt shape
            var canvasBuilder = new StringBuilder();
            foreach (var info in wordArtInfos)
            {
                canvasBuilder.AppendLine(
                    $"<canvas id=\"{info.Id}\" width=\"{info.Width}\" height=\"{info.Height}\" style=\"display:block;margin:10px 0;\"></canvas>");
            }

            // Build JavaScript to render gradients on the canvases
            var scriptBuilder = new StringBuilder();
            scriptBuilder.AppendLine("<script>");
            scriptBuilder.AppendLine("function drawWordArtGradients(){");
            foreach (var info in wordArtInfos)
            {
                scriptBuilder.AppendLine($"  var canvas = document.getElementById('{info.Id}');");
                scriptBuilder.AppendLine("  if (canvas && canvas.getContext){");
                scriptBuilder.AppendLine("    var ctx = canvas.getContext('2d');");
                scriptBuilder.AppendLine(
                    $"    var grad = ctx.createLinearGradient(0,0,canvas.width*Math.cos({info.Angle}*Math.PI/180),canvas.height*Math.sin({info.Angle}*Math.PI/180));");

                foreach (var stop in info.GradientStops)
                {
                    string hex = $"#{stop.Color.R:X2}{stop.Color.G:X2}{stop.Color.B:X2}";
                    scriptBuilder.AppendLine($"    grad.addColorStop({stop.Position}, '{hex}');");
                }

                scriptBuilder.AppendLine("    ctx.fillStyle = grad;");
                scriptBuilder.AppendLine("    ctx.fillRect(0,0,canvas.width,canvas.height);");
                scriptBuilder.AppendLine("  }");
                scriptBuilder.AppendLine("}");
            }
            scriptBuilder.AppendLine("window.onload = drawWordArtGradients;");
            scriptBuilder.AppendLine("</script>");

            // Insert canvases and script before </body>
            string insertion = canvasBuilder.ToString() + scriptBuilder.ToString();
            html = html.Replace("</body>", insertion + "\n</body>");

            // Write final HTML
            File.WriteAllText(outputPath, html);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}

// Helper class to store shape information
class WordArtInfo
{
    public string Id { get; set; } = string.Empty;
    public double Width { get; set; }
    public double Height { get; set; }
    public double Angle { get; set; } // In degrees
    public List<GradientStopInfo> GradientStops { get; set; } = new();
}

// Helper class for gradient stop details
class GradientStopInfo
{
    public Color Color { get; set; }
    public double Position { get; set; } // 0.0 to 1.0
}
