// Title: Export Excel WordArt with gradient fill to HTML and render it via Canvas and JavaScript using Aspose.Cells for .NET
// Common Searches: aspnet export excel wordart gradient fill to html using aspose.cells | replace wordart image tag with canvas element in aspose.cells html conversion | extract gradient stops from excel shape with aspose.cells c# | draw linear gradient on html canvas from excel wordart using javascript | convert excel worksheet to html preserving wordart gradients as canvas
// Tags: Aspose.Cells WordArt gradient HTML export | C# gradient stop extraction from Excel shape | HTML canvas placeholder for Excel WordArt | JavaScript linear gradient rendering from Excel data | Base64 image replacement with canvas in Aspose.Cells output

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace WordArtGradientHtmlExport
{
    // Helper class to store gradient information for a WordArt shape
    // The sample loads an Excel workbook, finds WordArt shapes that use gradient fills, extracts each gradient stop and the shape size, then saves the workbook as HTML with images embedded as Base64. It replaces the generated <img> tags for those WordArt shapes with uniquely‑identified <canvas> elements sized to the original shapes and injects JavaScript that draws the same linear gradient on each canvas using the extracted color stops. The final HTML preserves the visual appearance of WordArt gradients without relying on static images.
    class WordArtInfo
    {
        public string CanvasId { get; set; } = string.Empty;          // Unique canvas element id
        public int Width { get; set; }                                 // Shape width in pixels
        public int Height { get; set; }                                // Shape height in pixels
        public List<GradientStopInfo> Stops { get; set; } = new();    // Gradient stops
    }

    // Represents a single gradient stop (wrapper for Aspose.Cells.Drawing.GradientStop)
    class GradientStopInfo
    {
        public Color Color { get; set; } = Color.Empty;              // Stop color
        public double Position { get; set; }                          // Position (0.0 - 1.0)
    }

    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.html";

                // ---------- Load workbook ----------
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                Workbook workbook = new Workbook(inputPath);

                // ---------- Collect WordArt gradient data ----------
                List<WordArtInfo> wordArtInfos = new();
                int canvasCounter = 0;

                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Identify WordArt shapes via TextEffect property and gradient fill
                        if (shape.TextEffect != null && shape.Fill != null && shape.Fill.FillType == FillType.Gradient)
                        {
                            GradientFill gradientFill = shape.Fill.GradientFill;
                            if (gradientFill == null) continue;

                            List<GradientStopInfo> stops = new();
                            foreach (var stopObj in gradientFill.GradientStops)
                            {
                                // Use dynamic to access properties (avoids compile‑time binding issues)
                                dynamic aspStop = stopObj;
                                Color stopColor = aspStop.Color;
                                double position = aspStop.Position;

                                stops.Add(new GradientStopInfo
                                {
                                    Color = stopColor,
                                    Position = position
                                });
                            }

                            // Unique canvas id
                            string canvasId = $"wordart_canvas_{canvasCounter++}";

                            // Convert shape size from points to pixels (1 point ≈ 1.3333 pixels)
                            int widthPx = (int)Math.Round(shape.Width * 1.3333);
                            int heightPx = (int)Math.Round(shape.Height * 1.3333);

                            wordArtInfos.Add(new WordArtInfo
                            {
                                CanvasId = canvasId,
                                Width = widthPx,
                                Height = heightPx,
                                Stops = stops
                            });
                        }
                    }
                }

                // ---------- Convert workbook to HTML ----------
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportImagesAsBase64 = true,
                    ExportActiveWorksheetOnly = true
                };

                using (MemoryStream htmlStream = new MemoryStream())
                {
                    workbook.Save(htmlStream, htmlOptions);
                    htmlStream.Position = 0;
                    string htmlContent = new StreamReader(htmlStream).ReadToEnd();

                    // ---------- Replace WordArt images with canvas placeholders ----------
                    foreach (WordArtInfo info in wordArtInfos)
                    {
                        string canvasTag = $"<canvas id=\"{info.CanvasId}\" width=\"{info.Width}\" height=\"{info.Height}\"></canvas>";

                        int imgStart = htmlContent.IndexOf("<img", StringComparison.OrdinalIgnoreCase);
                        if (imgStart >= 0)
                        {
                            int imgEnd = htmlContent.IndexOf('>', imgStart);
                            if (imgEnd > imgStart)
                            {
                                htmlContent = htmlContent.Remove(imgStart, imgEnd - imgStart + 1)
                                                         .Insert(imgStart, canvasTag);
                            }
                        }
                    }

                    // ---------- Append JavaScript to draw gradients ----------
                    StringBuilder scriptBuilder = new StringBuilder();
                    scriptBuilder.AppendLine("<script type=\"text/javascript\">");
                    scriptBuilder.AppendLine("window.addEventListener('load', function() {");

                    foreach (WordArtInfo info in wordArtInfos)
                    {
                        scriptBuilder.AppendLine($"  var canvas = document.getElementById('{info.CanvasId}');");
                        scriptBuilder.AppendLine("  if (canvas && canvas.getContext) {");
                        scriptBuilder.AppendLine("    var ctx = canvas.getContext('2d');");
                        scriptBuilder.AppendLine("    var grad = ctx.createLinearGradient(0, 0, 0, canvas.height);");

                        foreach (GradientStopInfo stop in info.Stops)
                        {
                            string rgba = $"rgba({stop.Color.R}, {stop.Color.G}, {stop.Color.B}, {stop.Color.A / 255.0})";
                            scriptBuilder.AppendLine($"    grad.addColorStop({stop.Position.ToString(CultureInfo.InvariantCulture)}, '{rgba}');");
                        }

                        scriptBuilder.AppendLine("    ctx.fillStyle = grad;");
                        scriptBuilder.AppendLine("    ctx.fillRect(0, 0, canvas.width, canvas.height);");
                        scriptBuilder.AppendLine("  }");
                    }

                    scriptBuilder.AppendLine("});");
                    scriptBuilder.AppendLine("</script>");

                    // Insert script before closing </body>
                    int bodyCloseIdx = htmlContent.LastIndexOf("</body>", StringComparison.OrdinalIgnoreCase);
                    if (bodyCloseIdx >= 0)
                    {
                        htmlContent = htmlContent.Insert(bodyCloseIdx, scriptBuilder.ToString());
                    }
                    else
                    {
                        htmlContent += scriptBuilder.ToString();
                    }

                    // ---------- Save final HTML ----------
                    File.WriteAllText(outputPath, htmlContent, Encoding.UTF8);
                    Console.WriteLine($"HTML exported successfully to {outputPath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
