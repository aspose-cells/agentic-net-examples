// Title: Generate a single PNG that merges a column chart and its data table from an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a workbook, adds sample sales data, builds a column chart, and uses Aspose.Cells SheetRender to export the entire worksheet—including both the chart and the source data table—as one PNG image. | Demonstrate how to set ImageOrPrintOptions (e.g., OnePagePerSheet) so that the chart and its data table are rendered together on a single page when saving to PNG with Aspose.Cells.
// Common Searches: Aspose.Cells C# export worksheet with chart and data table to a single PNG file | how to render an Excel sheet containing a chart and its source table as one image using .NET | combine column chart image and data grid into one PNG with Aspose.Cells | C# generate composite image of Excel chart and table on one page | Aspose.Cells render chart and table together one-page PNG
// Tags: render worksheet chart to PNG Aspose.Cells | export column chart and data table as single image C# | Aspose.Cells composite chart image generation | ImageOrPrintOptions OnePagePerSheet .NET | SheetRender export Excel sheet to PNG

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The example creates a workbook with monthly sales data, adds a column chart, and uses Aspose.Cells to render the whole worksheet—including the chart and its data table—into a single PNG file.
class Program
{
    static void Main()
    {
        try
        {
            // 1. Create a workbook and fill it with sample data
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];

            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");

            string[] months = { "Jan", "Feb", "Mar", "Apr", "May" };
            double[] sales = { 1200, 1500, 1800, 1300, 1700 };

            for (int i = 0; i < months.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(months[i]);   // Column A
                sheet.Cells[i + 1, 1].PutValue(sales[i]);   // Column B
            }

            // 2. Add a column chart based on the data
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            var chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B6", true);
            // chart.NSeries[0].CategoryData = "A2:A6"; // optional category data
            chart.Title.Text = "Monthly Sales";

            // 3. Export the worksheet (including the chart and data table) to a PNG image
            var imgOptions = new ImageOrPrintOptions
            {
                // Default image format is PNG; no need to set explicitly
                OnePagePerSheet = true
            };

            string outputPath = "CompositeChart.png";

            // Ensure the output directory exists (handle possible null directory)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Render the sheet to an image file
            var sheetRender = new SheetRender(sheet, imgOptions);
            sheetRender.ToImage(0, outputPath);

            Console.WriteLine($"Composite chart saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
