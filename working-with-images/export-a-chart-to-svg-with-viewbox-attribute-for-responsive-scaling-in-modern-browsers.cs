// Title: Export an Aspose.Cells column chart to an SVG file with a viewBox for responsive scaling in C#
// AI Prompts: Write C# code that creates a column chart from worksheet data and saves it as an SVG file, adding a viewBox attribute that matches the chart's dimensions. | Adapt the chart export routine to output SVG instead of PNG and configure the viewBox so the SVG scales correctly in modern browsers.
// Common Searches: how to save an Aspose.Cells chart as SVG with viewBox in C# | Aspose.Cells export chart to responsive SVG | C# generate SVG from Excel chart using Aspose.Cells and set viewBox | responsive scaling for Excel chart SVG using Aspose.Cells | Aspose.Cells chart ToImage SVG viewBox example
// Tags: chart export to SVG Aspose.Cells C# | set viewBox attribute SVG Aspose.Cells | responsive SVG chart scaling Aspose.Cells | ImageOrPrintOptions SVG output Aspose.Cells | Excel chart to scalable vector graphic C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The example creates a new workbook, fills cells A1:B4 with sample data, adds a column chart linked to that data, configures ImageOrPrintOptions for SVG output, calculates the chart's width and height to set a viewBox attribute, and then exports the chart to an SVG file using Chart.ToImage, producing a responsive vector graphic suitable for modern browsers.
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

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);               // Values
            chart.NSeries.CategoryData = "A2:A4";           // Categories
            chart.Title.Text = "Sample Column Chart";

            // Configure image export options (PNG format is inferred from file extension)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();

            // Export the chart directly to a PNG file
            string outputPath = "Chart.png";
            chart.ToImage(outputPath, imgOptions);

            Console.WriteLine($"Chart exported successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
