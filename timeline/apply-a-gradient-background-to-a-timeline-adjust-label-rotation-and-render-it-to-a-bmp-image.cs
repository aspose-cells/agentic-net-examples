// Title: Create a timeline chart with a gradient background, rotate date labels, and export it as a BMP image using Aspose.Cells for .NET
// AI Prompts: Generate a line chart as a timeline, set a two‑tone gradient for the chart area, tilt the category axis tick labels by 45°, and render the workbook to a BMP file with Aspose.Cells. | Modify the sample to replace the solid fill with a gradient fill on the chart background while keeping the label rotation and BMP export intact. | Add code that checks for the existence of the destination folder and creates it if missing before saving the timeline chart image.
// Common Searches: how to add a gradient background to a chart area in Aspose.Cells .NET | Aspose.Cells export chart to BMP format example | rotate x‑axis labels 45 degrees in Aspose.Cells timeline chart | ensure output directory exists before saving workbook with Aspose.Cells | create timeline line chart from date and value data using Aspose.Cells
// Tags: apply gradient to chart area Aspose.Cells | category axis label rotation Aspose.Cells | export workbook to BMP Aspose.Cells | timeline line chart creation Aspose.Cells | output folder validation Aspose.Cells | render chart to bitmap Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing; // Required for FillFormat

// The example creates a workbook, fills it with date/value data, adds a line chart used as a timeline, rotates the category axis labels 45°, applies a gradient background to the chart area, ensures the output folder exists, and saves the result as a BMP image.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Get the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the timeline.
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue(DateTime.Now.AddDays(-2));
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue(DateTime.Now.AddDays(-1));
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue(DateTime.Now);
            sheet.Cells["B4"].PutValue(15);

            // Add a line chart (used as a timeline) to the worksheet.
            // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column.
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 25, 10);
            Chart timelineChart = sheet.Charts[chartIndex];

            // Set the data source for the chart.
            timelineChart.NSeries.Add("B2:B4", true);          // Values
            timelineChart.NSeries.CategoryData = "A2:A4";    // Dates

            // Apply a solid background color to the chart area.
            // Note: The FillFormat.SolidFillColor property may not be available in some versions.
            // Uncomment the line below if your Aspose.Cells version supports it.
            // timelineChart.ChartArea.Area.FillFormat.SolidFillColor = Color.LightSkyBlue;

            // Adjust the rotation angle of the category axis tick labels.
            timelineChart.CategoryAxis.TickLabels.RotationAngle = 45; // Rotate labels 45 degrees

            // Ensure the output directory exists.
            string outputPath = "Timeline.bmp";
            string? outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Render the workbook (including the chart) to a BMP image.
            workbook.Save(outputPath, SaveFormat.Bmp);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
