// Title: Generate and Export a Column Chart as SVG Using Aspose.Cells in C#
// AI Prompts: Write C# code that creates a column chart from worksheet data and saves it directly to an SVG file with Aspose.Cells. | Implement a method that accepts a Worksheet object and returns the rendered chart as an SVG byte array using Aspose.Cells ImageOrPrintOptions. | Adjust the SVG export to specify a custom canvas size and apply a transparent background for the chart rendered by Aspose.Cells.
// Common Searches: how to save an Aspose.Cells chart as an SVG file in C# | Aspose.Cells export column chart to scalable vector graphic example | C# code for rendering Excel chart to SVG using ImageOrPrintOptions | set custom SVG canvas dimensions when exporting chart with Aspose.Cells | retrieve chart SVG bytes from Aspose.Cells without writing to disk
// Tags: Aspose.Cells column chart SVG rendering | C# ImageOrPrintOptions SaveFormat Svg usage | chart.ToImage method for SVG output | Excel worksheet data to scalable vector graphic | custom SVG canvas size Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The example creates a new workbook, fills cells A1:B4 with month and sales data, adds a column chart, sets its title and data series, configures ImageOrPrintOptions to use the SVG save format, renders the chart to a memory stream, and writes the SVG bytes to a file named "Chart.svg".
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            using (Workbook workbook = new Workbook())
            {
                // Access the first worksheet
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
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
                Chart chart = sheet.Charts[chartIndex];

                // Set chart title
                chart.Title.Text = "Monthly Sales";

                // Define the data range for the series and categories
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Configure image options to export as SVG
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    SaveFormat = SaveFormat.Svg,
                    OnePagePerSheet = true
                };

                // Export the chart to an SVG file
                using (MemoryStream ms = new MemoryStream())
                {
                    chart.ToImage(ms, imgOptions);
                    File.WriteAllBytes("Chart.svg", ms.ToArray());
                }
            }
        }
        catch (Exception ex)
        {
            // Log or display the error
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
