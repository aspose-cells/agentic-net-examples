// Title: Generate an Aspose.Cells column chart in C#, export it as PNG, embed it as a base64 image in an HTML page, and add interactive zoom/pan with svg-pan-zoom
// AI Prompts: Write C# code that creates a workbook, adds sample data, builds a column chart, renders the chart to a PNG memory stream, converts the stream to a base64 string, and generates an HTML file that displays the image via a data URI. | Modify the program to render the chart as SVG instead of PNG, embed the SVG in the HTML page, and initialize the svg-pan-zoom library so the chart can be zoomed and panned in the browser. | Add CSS to define a fixed-size container for the chart image and ensure the embedded image scales correctly while preserving zoom and pan functionality.
// Common Searches: Aspose.Cells C# export chart to PNG base64 for web page | embed Aspose.Cells chart in HTML with zoom and pan | C# generate HTML file with chart image using memory stream | use svg-pan-zoom with Aspose.Cells chart output | how to render Aspose.Cells chart as SVG for interactive web display
// Tags: Aspose.Cells chart PNG export C# | C# base64 image embedding in HTML | svg-pan-zoom JavaScript library usage | interactive zoom and pan for web charts | Aspose.Cells render chart as SVG C#

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The example creates a workbook, populates it with sample data, adds a column chart, renders the chart to a PNG image in a memory stream, converts the image to a base64 data URI, builds an HTML page that embeds the PNG, loads the svg-pan-zoom script for potential interactive zoom/pan, and saves the page to disk.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(130);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);               // Values
            chart.NSeries.CategoryData = "A2:A4";           // Categories
            chart.Title.Text = "Monthly Sales";

            // Set image options for PNG export (default format is PNG, so ImageFormat is omitted)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true
            };

            // Render the chart into a memory stream
            using (MemoryStream imgStream = new MemoryStream())
            {
                chart.ToImage(imgStream, imgOptions);
                byte[] imgBytes = imgStream.ToArray();
                string base64Img = Convert.ToBase64String(imgBytes);

                // Build an HTML page that embeds the PNG image
                string html = $@"
<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8'>
    <title>Chart with Zoom & Pan</title>
    <script src='https://cdnjs.cloudflare.com/ajax/libs/svg-pan-zoom/3.6.1/svg-pan-zoom.min.js'></script>
    <style>
        #imgContainer {{
            width: 800px;
            height: 600px;
            border: 1px solid #ccc;
            overflow: hidden;
            display: flex;
            align-items: center;
            justify-content: center;
        }}
        img {{
            max-width: 100%;
            max-height: 100%;
        }}
    </style>
</head>
<body>
    <div id='imgContainer'>
        <img src='data:image/png;base64,{base64Img}' alt='Chart' />
    </div>
</body>
</html>";

                // Save the HTML file to disk
                string outputPath = "ChartWithZoom.html";
                File.WriteAllText(outputPath, html, Encoding.UTF8);
                Console.WriteLine($"HTML file saved to: {Path.GetFullPath(outputPath)}");
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
