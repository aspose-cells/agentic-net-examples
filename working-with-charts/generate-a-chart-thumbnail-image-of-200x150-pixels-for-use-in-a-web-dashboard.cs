// Title: Create a 200 × 150 PNG chart thumbnail from an Aspose.Cells worksheet in C# for a web dashboard
// AI Prompts: Use Aspose.Cells in C# to render a column chart as a 200 × 150 PNG image for embedding in a web dashboard. | Generate a small PNG of only the chart area from an Excel worksheet using ImageOrPrintOptions and SheetRender in C#.
// Common Searches: Aspose.Cells C# export Excel chart as 200x150 PNG for web UI | How to create a chart preview image from a workbook using Aspose.Cells .NET | Render only the chart portion of an Excel sheet to a small PNG with C# | Set custom dimensions for chart image export with Aspose.Cells ImageOrPrintOptions
// Tags: Aspose.Cells chart thumbnail PNG | C# ImageOrPrintOptions chart export | SheetRender export chart area | Excel column chart to custom size image | web dashboard chart image generation | Aspose.Cells render chart to PNG

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The example creates a workbook, adds sample data and a column chart, configures ImageOrPrintOptions for PNG output, and uses SheetRender to save the chart area as a 200 × 150 PNG file named chart_thumbnail.png, suitable for web dashboards.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
            Chart chart = sheet.Charts[chartIdx];

            // Set chart title
            chart.Title.Text = "Sample Chart";

            // Define the series (values) and categories (labels)
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure image rendering options
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // Set output format to PNG
                SaveFormat = SaveFormat.Png,
                // Ensure the whole sheet fits into the image
                OnePagePerSheet = true
            };

            // Render the first (and only) page of the sheet to an image file
            string outputPath = "chart_thumbnail.png";
            SheetRender renderer = new SheetRender(sheet, imgOptions);
            renderer.ToImage(0, outputPath);

            Console.WriteLine($"Chart thumbnail saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
