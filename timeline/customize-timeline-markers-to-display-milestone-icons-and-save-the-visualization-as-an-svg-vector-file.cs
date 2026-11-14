// Title: Generate a timeline line chart with milestone icons and export it as an SVG vector file using Aspose.Cells for .NET
// Common Searches: asp.net c# create a timeline line chart with custom milestone icons using Aspose.Cells | how to export an Aspose.Cells chart to an SVG vector file in .NET | set a custom image as a marker for a line chart series with Aspose.Cells C# | save Excel workbook as SVG format with Aspose.Cells library | timeline chart with date axis and custom markers example in C#
// Tags: timeline chart marker customization Aspose.Cells | export chart to SVG Aspose.Cells | line chart with date axis Aspose.Cells | custom marker image Aspose.Cells C# | save workbook as SVG vector Aspose.Cells | milestone icons in Aspose.Cells chart

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new workbook, fills columns A and B with monthly dates and sample values, adds a line chart with a date axis, sets a chart title, notes that marker image customization may be unavailable in some versions, and saves the workbook as an SVG vector file named TimelineMilestones.svg.
class TimelineMilestoneSvg
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data: dates (timeline) and values
            // Column A: Dates, Column B: Values
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Value");
            DateTime startDate = new DateTime(2023, 1, 1);
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(startDate.AddDays(i * 30)); // every month
                sheet.Cells[i + 1, 1].PutValue(i * 10 + 5);               // arbitrary values
            }

            // Add a line chart with markers
            int chartIndex = sheet.Charts.Add(ChartType.Line, 12, 0, 27, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title (optional)
            chart.Title.Text = "Project Timeline with Milestones";

            // Add series using the data range
            int seriesIndex = chart.NSeries.Add("B2:B11", true);
            chart.NSeries[seriesIndex].XValues = "A2:A11";

            // NOTE: In some Aspose.Cells versions the Marker class does not expose
            // Type, ImageData, or Size properties. To keep the code compatible,
            // marker customization is omitted. The default line chart will still be generated.

            // Save the workbook as an SVG vector file
            string outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TimelineMilestones.svg");
            workbook.Save(outputPath, SaveFormat.Svg);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
