// Title: How to set a custom subtitle on a chart and read back the title text using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# how to add a subtitle to a chart title | read chart title after setting subtitle using Aspose.Cells | example of column chart with custom subtitle in .NET | verify chart title text contains subtitle Aspose.Cells | include line break in chart title with Aspose.Cells C#
// Tags: custom chart subtitle Aspose.Cells C# | retrieve chart title text Aspose.Cells | generate column chart Aspose.Cells | newline in chart title Aspose.Cells | validate chart title update Aspose.Cells | save workbook with chart subtitle Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new workbook, adds a column chart, sets the chart's Title.Text to include a newline separating the main title and a custom subtitle, reads back the Title.Text to confirm the change, prints it to the console, and saves the workbook as ChartSubtitle.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data for the chart
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set the main title (including a subtitle using a line break)
            chart.Title.Text = "Sample Chart\nCustom subtitle text";

            // Verify the title (which now contains the subtitle)
            string title = chart.Title.Text;
            Console.WriteLine("Chart Title (with subtitle): " + title);

            // Save the workbook (optional, just to persist the chart)
            workbook.Save("ChartSubtitle.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
