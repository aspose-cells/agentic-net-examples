// Title: Generate a column chart with Aspose.Cells, export it as SVG, and embed it in an HTML page with interactive zoom and pan using svg-pan-zoom (C#)
// Common Searches: how to export an Aspose.Cells chart to SVG in C# | embed Aspose.Cells generated SVG chart in a web page with zoom and pan | use svg-pan-zoom to add interactive zoom to an SVG chart created by Aspose.Cells | C# code to create a column chart, save as SVG, and display with pan/zoom in HTML | integrate svg-pan-zoom library with Aspose.Cells SVG output for interactive web visualization
// Tags: Aspose.Cells chart SVG export C# | svg-pan-zoom library usage with generated SVG | interactive zoom and pan for SVG charts in web page | column chart SVG creation using Aspose.Cells | C# HTML embedding of SVG with pan/zoom controls

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;   // Required for ImageOrPrintOptions

// The example creates a workbook, adds sample data, builds a column chart, and exports the chart to an SVG file using Aspose.Cells. It then reads the SVG content, constructs an HTML page that embeds the SVG inside a container, and initializes the svg-pan-zoom library to provide interactive zoom and pan functionality before saving the HTML file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);               // Values
            chart.NSeries.CategoryData = "A2:A4";           // Categories
            chart.Title.Text = "Sample Column Chart";

            // Export the chart as an SVG file
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                SaveFormat = SaveFormat.Svg,
                OnePagePerSheet = true
            };
            string svgFilePath = "chart.svg";

            try
            {
                chart.ToImage(svgFilePath, imgOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to export chart to SVG: {ex.Message}");
                return;
            }

            // Ensure the SVG file was created before reading
            if (!File.Exists(svgFilePath))
                throw new FileNotFoundException("SVG file was not generated.", svgFilePath);

            // Read the generated SVG content
            string svgContent = File.ReadAllText(svgFilePath);

            // Build an HTML page that embeds the SVG and enables zoom/pan via svg-pan-zoom
            string htmlContent = $@"
<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8'>
    <title>SVG Chart with Zoom & Pan</title>
    <script src='https://cdnjs.cloudflare.com/ajax/libs/svg-pan-zoom/3.6.1/svg-pan-zoom.min.js'></script>
    <style>
        #svgContainer {{
            width: 800px;
            height: 600px;
            border: 1px solid #ccc;
            overflow: hidden;
        }}
    </style>
</head>
<body>
    <div id='svgContainer'>
        {svgContent}
    </div>
    <script>
        // Initialize svg-pan-zoom on the loaded SVG element
        var svgElem = document.querySelector('#svgContainer svg');
        if (svgElem) {{
            svgPanZoom(svgElem, {{
                zoomEnabled: true,
                controlIconsEnabled: true,
                fit: true,
                center: true
            }});
        }}
    </script>
</body>
</html>";

            // Save the HTML page to disk
            File.WriteAllText("ChartWithZoom.html", htmlContent);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
