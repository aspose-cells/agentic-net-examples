// Title: Create a timeline chart from Excel data and export it as a PNG image with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# generate line chart from date column and save as PNG | How to render an Excel chart to PNG using Aspose.Cells in a .NET console application | Export timeline chart created from worksheet data to PNG with Aspose.Cells | C# code to create a chart from dates and values and output PNG using Aspose.Cells | Save Aspose.Cells chart as an image file without opening Excel
// Tags: Aspose.Cells chart rendering to PNG | C# timeline chart creation Aspose.Cells | ImageOrPrintOptions PNG output Aspose.Cells | line chart from date series Aspose.Cells | default style chart generation Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The example creates a new workbook, fills column A with sequential dates and column B with numeric values, adds a line chart that serves as a timeline, configures ImageOrPrintOptions for PNG format, and renders the chart directly to a file named timeline.png, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Populate sample data: dates in column A, values in column B
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Value");

            DateTime startDate = new DateTime(2023, 1, 1);
            for (int i = 0; i < 10; i++)
            {
                // Column A: dates
                sheet.Cells[i + 1, 0].PutValue(startDate.AddDays(i));
                // Column B: numeric values
                sheet.Cells[i + 1, 1].PutValue(i * 10 + 5);
            }

            // Add a line chart (used as a substitute for timeline)
            int chartIndex = sheet.Charts.Add(ChartType.Line, 12, 0, 30, 10);
            var chart = sheet.Charts[chartIndex];

            // Set the series values (Y) and categories (X)
            chart.NSeries.Add("B2:B11", true);          // Values range
            chart.NSeries.CategoryData = "A2:A11";      // Dates range

            // Configure image options for PNG output
            var imgOptions = new ImageOrPrintOptions();
            imgOptions.SaveFormat = SaveFormat.Png;

            // Render the chart directly to a PNG file
            chart.ToImage("timeline.png", imgOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
