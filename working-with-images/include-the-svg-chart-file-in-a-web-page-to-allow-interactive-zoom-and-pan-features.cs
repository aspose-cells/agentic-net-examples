// Title: Export Aspose.Cells Chart to SVG and Add Interactive Zoom/Pan in HTML (C#)
// Description: Creates a workbook, builds a column chart, renders it to an SVG file with Aspose.Cells, and generates an HTML page that embeds the SVG and enables zoom and pan using the svg-pan-zoom JavaScript library. The solution is responsive and works in any modern browser.
// Keywords: Aspose.Cells SVG export | C# chart to SVG | interactive SVG chart | svg-pan-zoom integration | responsive chart HTML | .NET Excel chart rendering | embed SVG in web page
// Common Searches: export Aspose.Cells chart as SVG C# | add zoom and pan to SVG chart in HTML | svg-pan-zoom with Aspose.Cells output | responsive SVG chart from Excel data .NET | C# generate interactive SVG dashboard
// Developer Intent: Generate an SVG chart from an Excel workbook and embed it in a web page with client‑side zoom and pan capabilities.
// Use Cases: Build a lightweight, device‑agnostic sales dashboard that scales on mobile and desktop. | Create automated HTML reports where charts can be examined in detail without reloading the page. | Integrate Excel‑driven visualizations into existing web portals while preserving interactivity.
// AI Prompts: Show C# code that uses Aspose.Cells to export a chart to SVG and embeds it in an HTML file with svg-pan-zoom for zoom/pan. | Explain how to customize initial zoom level, control icons, and fit behavior of svg-pan-zoom in the generated page. | Provide an ASP.NET Core controller example that streams the SVG and HTML content to the browser while keeping interactive features.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// Creates a workbook, builds a column chart, renders it to an SVG file with Aspose.Cells, and generates an HTML page that embeds the SVG and enables zoom and pan using the svg-pan-zoom JavaScript library. The solution is responsive and works in any modern browser.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Month");
            worksheet.Cells["A2"].PutValue("Jan");
            worksheet.Cells["A3"].PutValue("Feb");
            worksheet.Cells["A4"].PutValue("Mar");
            worksheet.Cells["B1"].PutValue("Sales");
            worksheet.Cells["B2"].PutValue(120);
            worksheet.Cells["B3"].PutValue(150);
            worksheet.Cells["B4"].PutValue(180);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = worksheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";      // Categories
            chart.Title.Text = "Quarterly Sales";

            // Configure SVG rendering options
            SvgImageOptions svgOptions = new SvgImageOptions
            {
                FitToViewPort = true // Make SVG fit the viewport (responsive)
                // No need to set ImageFormat; SvgImageOptions is fixed to SVG
            };

            // Render the chart to an SVG file
            string svgFilePath = "chart.svg";
            chart.ToImage(svgFilePath, svgOptions);

            // Ensure the SVG file was created before embedding
            string svgContent = File.Exists(svgFilePath) ? File.ReadAllText(svgFilePath) : string.Empty;

            // Build an HTML page that embeds the SVG and enables zoom/pan via svg-pan-zoom library
            string htmlContent = $@"
<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8'>
    <title>Interactive SVG Chart</title>
    <script src='https://cdnjs.cloudflare.com/ajax/libs/svg-pan-zoom/3.6.1/svg-pan-zoom.min.js'></script>
    <style>
        #svgContainer {{
            width: 100%;
            height: 80vh;
            border: 1px solid #ccc;
        }}
        svg {{
            width: 100%;
            height: 100%;
        }}
    </style>
</head>
<body>
    <div id='svgContainer'>
        {svgContent}
    </div>
    <script>
        // Initialize pan and zoom functionality
        var panZoom = svgPanZoom('#svgContainer svg', {{
            zoomEnabled: true,
            controlIconsEnabled: true,
            fit: true,
            center: true
        }});
    </script>
</body>
</html>";

            // Save the HTML file
            File.WriteAllText("chart.html", htmlContent);

            Console.WriteLine("SVG chart generated and embedded in interactive HTML page.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
