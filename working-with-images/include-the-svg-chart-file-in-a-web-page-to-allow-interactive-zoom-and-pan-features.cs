// Title: Export Aspose.Cells Chart to SVG and Add Interactive Zoom/Pan with svg-pan-zoom (C#)
// Description: Creates a workbook, builds a column chart, renders it to an SVG file with a viewBox, and generates an HTML page that embeds the SVG and enables client‑side zoom and pan using the svg‑pan‑zoom library.
// Keywords: Aspose.Cells SVG export | C# chart to SVG | interactive SVG zoom | svg-pan-zoom integration | embed SVG in HTML | chart viewBox | client‑side pan and zoom | web dashboard charts
// Common Searches: Aspose.Cells export chart as SVG with viewBox | How to add zoom and pan to SVG chart in a web page | C# example for svg-pan-zoom with Aspose.Cells | Render Excel chart to interactive SVG | Embedding Aspose.Cells SVG in HTML
// Developer Intent: Generate an SVG version of an Aspose.Cells chart and embed it in a web page that supports interactive zoom and pan.
// Use Cases: Display sales or KPI charts on a responsive dashboard where users can explore details by zooming and panning. | Automate creation of static HTML reports that include scalable, interactive SVG graphics. | Integrate Aspose.Cells chart rendering into ASP.NET Core endpoints for on‑the‑fly SVG generation.
// AI Prompts: Write C# code that uses Aspose.Cells to export a chart to SVG with a viewBox and creates an HTML file that loads svg-pan-zoom for interactive zoom/pan. | Show how to customize the generated HTML to set the initial zoom level, toggle control icons, and define pan boundaries for the embedded SVG. | Explain how to serve the SVG and HTML files from an ASP.NET Core controller so the chart is rendered dynamically per request.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// Creates a workbook, builds a column chart, renders it to an SVG file with a viewBox, and generates an HTML page that embeds the SVG and enables client‑side zoom and pan using the svg‑pan‑zoom library.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and populate it with sample data
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            worksheet.Cells["A1"].PutValue("Month");
            worksheet.Cells["B1"].PutValue("Sales");
            worksheet.Cells["A2"].PutValue("Jan");
            worksheet.Cells["A3"].PutValue("Feb");
            worksheet.Cells["A4"].PutValue("Mar");
            worksheet.Cells["B2"].PutValue(120);
            worksheet.Cells["B3"].PutValue(150);
            worksheet.Cells["B4"].PutValue(180);

            // Add a column chart based on the data
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";
            chart.Title.Text = "Quarterly Sales";

            // Render the chart to an SVG file
            string svgFile = "chart.svg";
            try
            {
                SvgImageOptions svgOptions = new SvgImageOptions
                {
                    FitToViewPort = true // ensures a viewBox is generated
                };
                chart.ToImage(svgFile, svgOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error rendering SVG: {ex.Message}");
                return;
            }

            // Load the generated SVG content (ensure the file exists)
            string svgContent = File.Exists(svgFile) ? File.ReadAllText(svgFile) : string.Empty;

            // Build an HTML page that embeds the SVG and enables zoom/pan via svg-pan-zoom
            string htmlTemplate = @"<!DOCTYPE html>
<html>
<head>
    <meta charset=""utf-8"" />
    <title>Interactive SVG Chart</title>
    <style>
        body {{ margin:0; padding:0; overflow:hidden; }}
        #svgContainer {{ width:100vw; height:100vh; }}
    </style>
    <!-- svg-pan-zoom library from CDN -->
    <script src=""https://cdnjs.cloudflare.com/ajax/libs/svg-pan-zoom/3.6.1/svg-pan-zoom.min.js""></script>
</head>
<body>
    <div id=""svgContainer"">
        {0}
    </div>
    <script>
        document.addEventListener('DOMContentLoaded', function () {{
            var svg = document.querySelector('#svgContainer svg');
            if (svg) {{
                svgPanZoom(svg, {{
                    zoomEnabled: true,
                    controlIconsEnabled: true,
                    fit: true,
                    center: true
                }});
            }}
        }});
    </script>
</body>
</html>";
            string html = string.Format(htmlTemplate, svgContent);

            // Save the HTML page
            string htmlFile = "chart.html";
            try
            {
                File.WriteAllText(htmlFile, html);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error writing HTML file: {ex.Message}");
                return;
            }

            Console.WriteLine($"SVG chart generated: {svgFile}");
            Console.WriteLine($"HTML page with interactive zoom/pan generated: {htmlFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
