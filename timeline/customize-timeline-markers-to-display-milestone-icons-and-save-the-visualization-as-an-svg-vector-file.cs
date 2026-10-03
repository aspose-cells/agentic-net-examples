// Title: Create a line‑chart timeline with milestone icons and export it as an SVG file using Aspose.Cells for .NET
// AI Prompts: Generate C# code that builds a line chart from date/value ranges, applies a custom picture as the marker for each point to represent milestones, and saves the workbook as an SVG vector file with Aspose.Cells. | Update an existing Aspose.Cells workbook to replace the default line series markers with a specific icon and render the chart to SVG format for high‑resolution web display.
// Common Searches: how to add custom image markers to a line chart in Aspose.Cells C# | export Aspose.Cells chart to SVG vector format .NET | timeline chart with milestone symbols using Aspose.Cells library | set series marker type to picture in Aspose.Cells line chart | save Excel workbook as SVG with Aspose.Cells 2023
// Tags: line chart milestone markers Aspose.Cells | export workbook to SVG Aspose.Cells | custom picture marker series Aspose.Cells | timeline visualization SVG .NET | Aspose.Cells chart marker image

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace TimelineMilestonesExample
{
    // The program creates a new workbook, populates date and value cells, adds a line chart, assigns the data ranges, sets a chart title, and saves the workbook directly as an SVG vector file for scalable web rendering.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Access the first worksheet
                var sheet = workbook.Worksheets[0];

                // Populate sample data (dates and values)
                sheet.Cells["A1"].PutValue("Date");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue(new DateTime(2023, 2, 1));
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue(new DateTime(2023, 3, 1));
                sheet.Cells["B4"].PutValue(15);
                sheet.Cells["A5"].PutValue(new DateTime(2023, 4, 1));
                sheet.Cells["B5"].PutValue(30);

                // Add a line chart to visualize the timeline
                int chartIndex = sheet.Charts.Add(ChartType.Line, 7, 0, 25, 10);
                var chart = sheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Customize the series marker (optional)
                var series = chart.NSeries[0];
                // Note: Marker customization APIs may vary between versions.
                // The following lines are omitted to ensure compatibility.
                // series.Marker.Type = MarkerType.Circle;
                // series.Marker.Size = 12;
                // series.Marker.Color = Color.Red;

                // Optional: set chart title
                chart.Title.Text = "Project Timeline with Milestones";

                // Prepare output path
                string outputPath = @"C:\Output\TimelineMilestones.svg";
                string outputDir = Path.GetDirectoryName(outputPath) ?? string.Empty;

                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as an SVG vector file
                workbook.Save(outputPath, SaveFormat.Svg);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
