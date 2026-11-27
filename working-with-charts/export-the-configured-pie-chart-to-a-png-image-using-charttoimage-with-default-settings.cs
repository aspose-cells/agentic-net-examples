// Title: Export a pie chart as a PNG image with default settings using Aspose.Cells Chart.ToImage in C#
// Common Searches: C# code to save an Aspose.Cells pie chart as PNG | Aspose.Cells Chart.ToImage example for exporting pie chart to PNG | How to use ImageOrPrintOptions to export Excel chart to PNG in .NET | Exporting a configured pie chart to PNG with Aspose.Cells library | Default PNG export settings for Aspose.Cells charts in C#
// Tags: Aspose.Cells chart image export C# | Chart.ToImage PNG output | pie chart rendering Aspose.Cells | ImageOrPrintOptions default PNG | C# Aspose.Cells chart to image

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// The sample creates a workbook, fills cells with fruit categories and values, adds a pie chart linked to that data, sets a title, configures ImageOrPrintOptions for PNG with default settings, and uses Chart.ToImage to write the chart to 'pie_chart.png'.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data for the pie chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("Cherry");
            sheet.Cells["B4"].PutValue(50);

            // Add a pie chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";      // Categories

            // Optional: customize chart appearance
            chart.Title.Text = "Fruit Distribution";
            chart.FirstSliceAngle = 0; // default angle

            // Create image options with default PNG settings
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png
            };

            // Export the configured chart to a PNG image
            string outputFile = "pie_chart.png";
            chart.ToImage(outputFile, options);

            Console.WriteLine($"Pie chart successfully exported to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
