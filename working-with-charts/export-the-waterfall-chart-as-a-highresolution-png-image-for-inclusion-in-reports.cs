// Title: Export a Waterfall chart to a high‑resolution PNG file using Aspose.Cells for .NET
// AI Prompts: Write C# code that builds a Waterfall chart in an Aspose.Cells workbook and saves it as a 300 DPI PNG image. | Show how to configure ImageOrPrintOptions.HorizontalResolution and VerticalResolution for high‑resolution chart image export in Aspose.Cells. | Adapt the sample to specify custom image width and height while keeping a 300 DPI resolution for the exported chart.
// Common Searches: Aspose.Cells export waterfall chart as 300 dpi PNG in C# | how to set DPI for chart image export with Aspose.Cells .NET | C# code to save Excel Waterfall chart to high‑resolution PNG using Aspose | ImageOrPrintOptions high resolution chart rendering Aspose.Cells | export Excel chart to PNG with specific resolution using Aspose.Cells
// Tags: waterfall chart export PNG Aspose.Cells | high‑resolution chart image Aspose.Cells .NET | ImageOrPrintOptions DPI setting | chart rendering to image C# | Aspose.Cells chart export options

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The example creates a workbook, fills it with data for a Waterfall chart, adds the chart, configures ImageOrPrintOptions to 300 DPI, and exports the chart as a high‑resolution PNG file named WaterfallChart.png.
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

            // Populate data for the waterfall chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");

            string[] categories = { "Start", "Revenue", "Cost", "Profit", "End" };
            double[] values = { 1000, 300, -200, 500, 1600 };

            for (int i = 0; i < categories.Length; i++)
            {
                sheet.Cells[i + 2, 0].PutValue(categories[i]); // Column A
                sheet.Cells[i + 2, 1].PutValue(values[i]);    // Column B
            }

            // Add a Waterfall chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 5, 0, 25, 10);
            var chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B6", true);
            chart.NSeries.CategoryData = "A2:A6";

            // Optional: set a chart title
            chart.Title.Text = "Waterfall Chart Example";

            // Configure image export options for high resolution (e.g., 300 DPI)
            var imgOptions = new ImageOrPrintOptions
            {
                // The format is inferred from the file extension, so ImageFormat is omitted
                HorizontalResolution = 300,
                VerticalResolution = 300
                // ImageWidth and ImageHeight are not supported in this version; they can be omitted
            };

            // Export the chart as a high‑resolution PNG image
            string outputPath = "WaterfallChart.png";
            chart.ToImage(outputPath, imgOptions);
            Console.WriteLine($"Chart exported successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
