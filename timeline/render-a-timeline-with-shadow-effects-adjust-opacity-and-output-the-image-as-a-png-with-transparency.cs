// Title: Render a timeline line chart with 30% transparent plot area and save it as a PNG with alpha channel using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells render line chart to PNG with transparent background | How to set plot area opacity for a chart in Aspose.Cells .NET | Export Excel timeline chart as PNG with alpha channel using Aspose.Cells | Aspose.Cells ImageOrPrintOptions transparent PNG example in C# | Create a chart with shadow effect and 30% transparency in Aspose.Cells
// Tags: render chart to transparent PNG Aspose.Cells | plot area transparency setting Aspose.Cells | timeline line chart export PNG C# | ImageOrPrintOptions transparent background Aspose.Cells | chart shadow effect Aspose.Cells | alpha channel PNG output Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace TimelineRenderDemo
{
    // The sample creates a workbook, fills it with date/value data, adds a line chart to represent a timeline, sets the plot area transparency to 30 %, configures ImageOrPrintOptions for PNG with a transparent background, and renders the first worksheet page to a PNG file named timeline_with_shadow.png.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Populate sample data for a timeline (dates and values)
                sheet.Cells["A1"].PutValue("Date");
                sheet.Cells["B1"].PutValue("Value");

                DateTime startDate = new DateTime(2023, 1, 1);
                for (int i = 0; i < 10; i++)
                {
                    sheet.Cells[i + 1, 0].PutValue(startDate.AddDays(i * 30)); // every month
                    sheet.Cells[i + 1, 1].PutValue(10 + i * 5);
                }

                // 3. Add a line chart (Timeline chart type is not available in older versions)
                int chartIndex = sheet.Charts.Add(ChartType.Line, 12, 0, 30, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B11", true);
                chart.NSeries.CategoryData = "A2:A11";

                // 4. Adjust opacity (transparency) of the chart's plot area (0 = opaque, 1 = fully transparent)
                // 30 % transparent
                chart.PlotArea.Area.FillFormat.Transparency = 0.3;

                // 5. Configure image rendering options for PNG with transparent background
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Png, // Aspose.Cells.Drawing.ImageType
                    Transparent = true,
                    HorizontalResolution = 300,
                    VerticalResolution = 300
                };

                // 6. Render the worksheet (first page) to a PNG file
                SheetRender renderer = new SheetRender(sheet, imgOptions);
                if (renderer.PageCount > 0)
                {
                    string outputPath = "timeline_with_shadow.png";
                    try
                    {
                        renderer.ToImage(0, outputPath);
                        Console.WriteLine($"Timeline image successfully saved to '{outputPath}'.");
                    }
                    catch (Exception renderEx)
                    {
                        Console.WriteLine("Error during rendering:");
                        Console.WriteLine(renderEx.Message);
                    }
                }
                else
                {
                    Console.WriteLine("No pages were generated for rendering.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during processing:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
