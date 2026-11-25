// Title: Create a 200 × 150 PNG thumbnail of a column chart using Aspose.Cells in C#
// Common Searches: how to export an Aspose.Cells chart as a 200x150 PNG image in C# | Aspose.Cells SetDesiredSize method to generate chart thumbnail for a web dashboard | render a column chart to a custom sized PNG with Aspose.Cells C# example | C# code to create a chart thumbnail image using Aspose.Cells ImageOrPrintOptions | generate a small chart image for dashboard using Aspose.Cells chart rendering
// Tags: Aspose.Cells chart rendering to PNG | ImageOrPrintOptions SetDesiredSize custom dimensions | export column chart as thumbnail C# | chart thumbnail generation for web dashboard | control chart image size Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace ChartThumbnailGenerator
{
    // The sample creates a workbook, fills cells A1:B4 with category and value data, adds a column chart, configures ImageOrPrintOptions to force a 200 × 150 PNG size without preserving aspect ratio, renders the chart to 'chart_thumbnail.png', and includes basic exception handling.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(30);

                // 3. Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // 4. Set the data source for the chart
                chart.NSeries.Add("B2:B4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";      // Categories

                // 5. Configure image rendering options
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = Aspose.Cells.Drawing.ImageType.Png,
                    // Ensure the whole chart fits into the desired dimensions without preserving aspect ratio
                    // (keepAspectRatio = false)
                };
                imgOptions.SetDesiredSize(200, 150, false);

                // 6. Render the chart to a PNG file with the specified thumbnail size
                string outputPath = "chart_thumbnail.png";
                chart.ToImage(outputPath, imgOptions);

                Console.WriteLine($"Chart thumbnail generated successfully: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("An error occurred while generating the chart thumbnail:");
                Console.Error.WriteLine(ex.Message);
            }
        }
    }
}
