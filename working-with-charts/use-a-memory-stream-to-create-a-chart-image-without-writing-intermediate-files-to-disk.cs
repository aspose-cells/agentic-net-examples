// Title: Create a PNG chart image from an Aspose.Cells workbook using a MemoryStream in C# (no temporary files)
// Common Searches: how to export an Aspose.Cells chart to a PNG stream in C# | render chart from workbook to memory without saving a file Aspose.Cells | Aspose.Cells C# generate chart image in a memory stream | save Aspose.Cells chart as PNG directly to a byte array | convert worksheet chart to image using SheetRender and MemoryStream
// Tags: Aspose.Cells chart to PNG via MemoryStream | C# SheetRender export chart image | in‑memory chart rendering with Aspose.Cells | column chart image generation Aspose.Cells | eliminate temporary files chart export C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace ChartImageToMemoryStream
{
    // The example creates a workbook, adds sample data and a column chart, configures PNG rendering options, and uses SheetRender to render the worksheet (including the chart) directly into a MemoryStream, avoiding any intermediate file writes. The resulting stream can be saved, sent over a network, or processed further.
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
                sheet.Cells["A5"].PutValue("D");
                sheet.Cells["B5"].PutValue(40);

                // 3. Add a column chart to the worksheet
                // Parameters: chart type, top row, left column, bottom row, right column
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Optional: give the chart a title
                chart.Title.Text = "Sample Column Chart";

                // 4. Prepare rendering options to output a PNG image
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Png,
                    OnePagePerSheet = true
                };

                // 5. Render the worksheet (which contains the chart) to a memory stream
                using (MemoryStream imageStream = new MemoryStream())
                {
                    try
                    {
                        SheetRender sheetRender = new SheetRender(sheet, renderOptions);
                        sheetRender.ToImage(0, imageStream);
                    }
                    catch (Exception renderEx)
                    {
                        Console.WriteLine("Error during rendering: " + renderEx.Message);
                        return;
                    }

                    if (imageStream.Length == 0)
                    {
                        Console.WriteLine("Failed to render chart image: stream is empty.");
                        return;
                    }

                    // 6. (Optional) Save the image to disk for verification.
                    // This step is only for demonstration and can be removed.
                    string outputPath = Path.Combine(Environment.CurrentDirectory, "ChartImage.png");
                    try
                    {
                        File.WriteAllBytes(outputPath, imageStream.ToArray());
                        Console.WriteLine($"Chart image successfully rendered to memory stream ({imageStream.Length} bytes).");
                        Console.WriteLine($"Image saved to: {outputPath}");
                    }
                    catch (Exception ioEx)
                    {
                        Console.WriteLine("Error writing image to disk: " + ioEx.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                // Output any unexpected errors
                Console.WriteLine("An error occurred while creating the chart image:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
