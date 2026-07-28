// Title: Export Aspose.Cells Chart to Transparent PNG in C# (.NET)
// Description: This C# example creates a workbook, adds sample data, builds a column chart, sets the chart area to transparent, enables transparent rendering via ImageOrPrintOptions, and saves the chart as a PNG with no background.
// Keywords: Aspose.Cells | C# | .NET | chart export | transparent PNG | ChartArea.BackgroundMode | ImageOrPrintOptions.Transparent | chart to image | transparent background | Aspose.Cells chart PNG
// Common Searches: Aspose.Cells export chart transparent PNG C# | How to make chart background transparent in Aspose.Cells | ImageOrPrintOptions Transparent true Aspose.Cells | ChartArea.BackgroundMode Transparent example | Save Aspose.Cells chart as PNG with no background
// Developer Intent: Generate a chart image with a transparent background using Aspose.Cells.
// Use Cases: Embedding chart PNGs over custom UI backgrounds in web applications. | Creating overlay graphics for dashboards without white borders. | Producing chart assets for PDF reports where the background should blend with the page color. | Automating batch export of multiple charts as transparent PNG files.
// AI Prompts: Write C# code that creates a pie chart with Aspose.Cells and exports it to a transparent PNG. | Explain the difference between ChartArea.BackgroundMode and ImageOrPrintOptions.Transparent when exporting charts. | Show how to export multiple Aspose.Cells charts to transparent PNG files in a loop. | Provide troubleshooting steps if the exported PNG still has a white background.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsTransparentChartExport
{
    // This C# example creates a workbook, adds sample data, builds a column chart, sets the chart area to transparent, enables transparent rendering via ImageOrPrintOptions, and saves the chart as a PNG with no background.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["A3"].PutValue("Orange");
                worksheet.Cells["A4"].PutValue("Banana");

                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(15);
                worksheet.Cells["B4"].PutValue(7);

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Make the chart area transparent (if supported)
                chart.ChartArea.BackgroundMode = BackgroundMode.Transparent;

                // Configure image export options for PNG with transparent background
                ImageOrPrintOptions options = new ImageOrPrintOptions
                {
                    // Default image format is PNG; no need to set explicitly
                    Transparent = true // Enable transparent background
                };

                // Export the chart to a PNG file using the configured options
                chart.ToImage("TransparentChart.png", options);

                Console.WriteLine("Chart exported to TransparentChart.png with transparent background.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
