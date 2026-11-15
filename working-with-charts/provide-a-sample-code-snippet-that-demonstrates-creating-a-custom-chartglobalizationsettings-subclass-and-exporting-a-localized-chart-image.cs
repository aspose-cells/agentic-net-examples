// Title: How to export a column chart as a PNG image using Aspose.Cells for .NET
// Common Searches: Aspose.Cells export worksheet chart to PNG file in C# | C# code to save a column chart as an image with Aspose.Cells | Using Chart.ToImage method to generate chart PNG in Aspose.Cells for .NET | Create and export a chart image from a workbook with Aspose.Cells | Save Excel chart as PNG using Aspose.Cells library in .NET
// Tags: export chart to PNG Aspose.Cells | Chart.ToImage method .NET | column chart image generation C# | Aspose.Cells workbook chart export | save Excel chart as image Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new Workbook, fills month and sales data, adds a column chart, defines its data range, and saves the chart as a PNG file using the Chart.ToImage overload that infers the format from the file extension.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Access the first worksheet
            Worksheet ws = wb.Worksheets[0];

            // Populate data for the chart
            ws.Cells["A1"].PutValue("Month");
            ws.Cells["B1"].PutValue("Sales");

            string[] months = { "Jan", "Feb", "Mar", "Apr", "May", "Jun" };
            double[] sales = { 1200.5, 1500.75, 1100.0, 1800.25, 1600.6, 1900.9 };

            for (int i = 0; i < months.Length; i++)
            {
                ws.Cells[i + 1, 0].PutValue(months[i]);
                ws.Cells[i + 1, 1].PutValue(sales[i]);
            }

            // Add a column chart to the worksheet
            int chartIndex = ws.Charts.Add(ChartType.Column, 7, 0, 20, 10);
            Chart chart = ws.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B7", true);
            chart.NSeries.CategoryData = "A2:A7";

            // Export the chart as a PNG image
            string outputPath = "LocalizedChart.png";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Use overload that infers format from file extension
            chart.ToImage(outputPath);
            Console.WriteLine($"Chart image saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
