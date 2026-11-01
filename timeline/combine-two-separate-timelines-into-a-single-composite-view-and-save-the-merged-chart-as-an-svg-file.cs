// Title: Create a composite timeline line chart from two date‑value series and export it as an SVG file using Aspose.Cells for .NET (C#)
// Common Searches: how to plot two separate date columns on one line chart with Aspose.Cells C# | Aspose.Cells C# export line chart as SVG image | combine multiple timeline series into a single chart using Aspose.Cells for .NET | C# code to add two data series to a line chart in Aspose.Cells workbook | save Aspose.Cells chart to SVG format programmatically
// Tags: Aspose.Cells line chart multiple series | export Aspose.Cells chart to SVG | C# composite timeline chart | Aspose.Cells chart rendering options | combine two data series Aspose.Cells | SVG output for Aspose.Cells charts

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The example creates a workbook with two separate date‑value series, adds a line chart that includes both series as a composite timeline, sets a chart title, and saves the chart as an SVG file named CompositeTimeline.svg using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "TimelineData";

            // ----- First timeline data -----
            // Header
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Value1");
            // Data rows
            sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue(new DateTime(2023, 2, 1));
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue(new DateTime(2023, 3, 1));
            sheet.Cells["B4"].PutValue(15);

            // ----- Second timeline data (placed in columns D and E) -----
            sheet.Cells["D1"].PutValue("Date");
            sheet.Cells["E1"].PutValue("Value2");
            sheet.Cells["D2"].PutValue(new DateTime(2023, 1, 15));
            sheet.Cells["E2"].PutValue(12);
            sheet.Cells["D3"].PutValue(new DateTime(2023, 2, 15));
            sheet.Cells["E3"].PutValue(18);
            sheet.Cells["D4"].PutValue(new DateTime(2023, 3, 15));
            sheet.Cells["E4"].PutValue(22);

            // Add a line chart that will hold both timelines
            // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // First series: values from B2:B4, X‑values from A2:A4
            int series1 = chart.NSeries.Add("B2:B4", true);
            chart.NSeries[series1].XValues = "A2:A4";

            // Second series: values from E2:E4, X‑values from D2:D4
            int series2 = chart.NSeries.Add("E2:E4", true);
            chart.NSeries[series2].XValues = "D2:D4";

            // Optional: set series names for clarity
            chart.NSeries[series1].Name = "Timeline 1";
            chart.NSeries[series2].Name = "Timeline 2";

            // Set chart title
            chart.Title.Text = "Composite Timeline View";

            // Export the chart as an SVG file
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // Use SaveFormat to specify SVG output (compatible with current API)
                SaveFormat = SaveFormat.Svg
            };
            chart.ToImage("CompositeTimeline.svg", imgOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
