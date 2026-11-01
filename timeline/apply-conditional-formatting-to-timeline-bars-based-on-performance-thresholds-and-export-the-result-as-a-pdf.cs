// Title: Create a timeline bar chart with performance-based conditional formatting and export it to PDF using Aspose.Cells for .NET
// Common Searches: how to color individual bars in an Aspose.Cells bar chart based on value | Aspose.Cells C# set conditional formatting for chart points | export a timeline style bar chart to PDF with Aspose.Cells | apply performance thresholds to bar colors in Aspose.Cells chart | C# Aspose.Cells change bar fill color programmatically
// Tags: conditional formatting chart points Aspose.Cells | timeline bar chart Aspose.Cells C# | set chart point fill color .NET | export workbook to PDF Aspose.Cells | performance threshold bar colors Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, fills it with dates and performance values, adds a bar chart styled as a timeline, colors each bar green, yellow, or red according to defined performance thresholds, configures the value axis, and saves the result as a PDF file.
class TimelineConditionalFormatting
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data: Date and Performance value
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Performance");

            DateTime startDate = new DateTime(2023, 1, 1);
            double[] performances = { 92, 67, 45, 78, 55, 83, 30 };
            for (int i = 0; i < performances.Length; i++)
            {
                // Column A: Date (increment by 7 days)
                sheet.Cells[i + 1, 0].PutValue(startDate.AddDays(i * 7));
                // Column B: Performance value
                sheet.Cells[i + 1, 1].PutValue(performances[i]);
            }

            // Add a bar chart (timeline style)
            int chartIndex = sheet.Charts.Add(ChartType.Bar, 1, 3, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set data source for the chart
            chart.NSeries.Add("B2:B8", true);          // Values (Y) axis
            chart.NSeries.CategoryData = "A2:A8";     // Category (X) axis

            // Set chart title
            chart.Title.Text = "Performance Timeline";

            // Apply conditional formatting to each data point based on thresholds
            // >=80 Green, >=50 Yellow, <50 Red
            for (int i = 0; i < performances.Length; i++)
            {
                ChartPoint point = chart.NSeries[0].Points[i];

                Color fillColor = performances[i] >= 80
                    ? Color.Green
                    : performances[i] >= 50 ? Color.Yellow : Color.Red;

                // Set the fill color for the data point
                point.Area.ForegroundColor = fillColor;
                // Ensure solid fill (no pattern property in current API)
                point.Area.BackgroundColor = fillColor;
            }

            // Adjust value axis (optional)
            chart.ValueAxis.MinValue = 0;
            chart.ValueAxis.MaxValue = 100;
            chart.ValueAxis.MajorUnit = 20;

            // Save the workbook as PDF
            workbook.Save("TimelineConditionalFormatting.pdf", SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
