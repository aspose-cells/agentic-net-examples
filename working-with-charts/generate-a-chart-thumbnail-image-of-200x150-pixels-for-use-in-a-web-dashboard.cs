// Title: Create a 200 × 150 px PNG chart thumbnail with Aspose.Cells for .NET (C#)
// Description: The sample builds a workbook, adds sample data, creates a column chart, and uses ImageOrPrintOptions.SetDesiredSize(200,150,false) to export the chart as a 200 × 150 pixel PNG file (chart_thumbnail.png) suitable for web‑dashboard widgets.
// Keywords: Aspose.Cells C# chart thumbnail | SetDesiredSize 200x150 | ImageOrPrintOptions export PNG | column chart preview Aspose | fixed size chart image .NET | web dashboard chart thumbnail | Aspose.Cells generate PNG preview
// Common Searches: Aspose.Cells export chart as 200x150 PNG | C# set chart image size without aspect ratio | How to create chart thumbnail for dashboard using Aspose.Cells | ImageOrPrintOptions SetDesiredSize example | Generate fixed‑size chart image in .NET
// Developer Intent: Generate a PNG thumbnail of a workbook chart with exact dimensions of 200 × 150 pixels using Aspose.Cells in C#.
// Use Cases: Display uniform chart previews in dashboard tiles. | Embed small chart images in email or PDF reports where layout consistency is required. | Serve lightweight chart thumbnails to mobile or low‑bandwidth web applications.
// AI Prompts: Show how to change the code to output a 300 × 200 JPEG thumbnail while preserving aspect ratio. | Provide a loop that saves every chart in a workbook as a 200 × 150 PNG thumbnail. | Add validation to ensure the chart data range is not empty before generating the thumbnail.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    // The sample builds a workbook, adds sample data, creates a column chart, and uses ImageOrPrintOptions.SetDesiredSize(200,150,false) to export the chart as a 200 × 150 pixel PNG file (chart_thumbnail.png) suitable for web‑dashboard widgets.
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
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["A3"].PutValue("Orange");
                sheet.Cells["A4"].PutValue("Banana");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(1200);
                sheet.Cells["B3"].PutValue(800);
                sheet.Cells["B4"].PutValue(1500);

                // Add a column chart (top‑left row, left column, bottom‑right row, right column)
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIdx];

                // Set the data range for the chart
                chart.SetChartDataRange("A1:B4", true);

                // Prepare image options – set the desired thumbnail size (200 × 150 px) without keeping aspect ratio
                ImageOrPrintOptions imgOpts = new ImageOrPrintOptions
                {
                    // The file extension determines the image format, so explicit ImageType assignment is optional
                };
                imgOpts.SetDesiredSize(200, 150, false); // width, height, keepAspectRatio

                // Export the chart as a PNG thumbnail
                string outputPath = "chart_thumbnail.png";
                chart.ToImage(outputPath, imgOpts);
                Console.WriteLine($"Chart thumbnail saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
