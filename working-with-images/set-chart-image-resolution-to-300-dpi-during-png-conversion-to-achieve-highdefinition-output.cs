// Title: Generate a 300 DPI PNG image of an Excel chart using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export chart to PNG with 300 DPI resolution | How to set ImageOrPrintOptions DPI for chart rendering in Aspose.Cells | Render Excel worksheet to high‑resolution PNG using Aspose.Cells .NET | C# create high‑definition chart image from .xlsx with Aspose.Cells | Set horizontal and vertical resolution when saving chart as PNG in Aspose.Cells
// Tags: chart rendering to PNG Aspose.Cells | ImageOrPrintOptions DPI configuration | export Excel chart high‑resolution image | worksheet page PNG 300 DPI Aspose.Cells | C# Aspose.Cells chart image generation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace ChartImageResolutionDemo
{
    // The sample program creates a workbook with a column chart if the source file is missing, loads the workbook, configures ImageOrPrintOptions with 300 DPI horizontal and vertical resolution, and uses Chart.ToImage to export the first chart as a high‑definition PNG. It also demonstrates rendering the entire worksheet to a 300 DPI PNG using SheetRender.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define input and output file names
                string inputFile = "SampleChart.xlsx";
                string outputChartImage = "Chart_300dpi.png";

                // Ensure the input workbook exists; if not, create a sample workbook with a chart
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file \"{inputFile}\" not found. Creating a sample workbook...");

                    // Create a new workbook and get the first worksheet
                    Workbook wbCreate = new Workbook();
                    Worksheet ws = wbCreate.Worksheets[0];

                    // Populate sample data for the chart
                    ws.Cells["A1"].PutValue("Category");
                    ws.Cells["B1"].PutValue("Value");
                    ws.Cells["A2"].PutValue("Product A");
                    ws.Cells["B2"].PutValue(120);
                    ws.Cells["A3"].PutValue("Product B");
                    ws.Cells["B3"].PutValue(150);
                    ws.Cells["A4"].PutValue("Product C");
                    ws.Cells["B4"].PutValue(90);
                    ws.Cells["A5"].PutValue("Product D");
                    ws.Cells["B5"].PutValue(200);

                    // Add a column chart
                    int chartIndex = ws.Charts.Add(ChartType.Column, 7, 0, 20, 10);
                    Chart chart = ws.Charts[chartIndex];
                    chart.Title.Text = "Sales by Product";
                    chart.NSeries.Add("B2:B5", true);
                    chart.NSeries.CategoryData = "A2:A5";

                    // Save the workbook for later use
                    wbCreate.Save(inputFile, SaveFormat.Xlsx);
                    Console.WriteLine($"Sample workbook saved as \"{inputFile}\".");
                }

                // Load the workbook containing the chart
                Workbook workbook = new Workbook(inputFile);
                Worksheet worksheet = workbook.Worksheets[0];

                // Find the first chart in the worksheet
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the worksheet. Exiting.");
                    return;
                }

                Chart targetChart = worksheet.Charts[0];
                Console.WriteLine($"Rendering chart \"{targetChart.Name}\" to PNG with 300 DPI...");

                // Configure image options for high‑definition output
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Png,
                    HorizontalResolution = 300, // 300 DPI horizontally
                    VerticalResolution = 300    // 300 DPI vertically
                };

                // Render the chart to a PNG file using the specified options
                targetChart.ToImage(outputChartImage, imgOptions);

                // Verify that the file was created
                if (File.Exists(outputChartImage))
                {
                    Console.WriteLine($"Chart image successfully saved as \"{outputChartImage}\" with 300 DPI.");
                }
                else
                {
                    Console.WriteLine("Failed to create the chart image file.");
                }

                // OPTIONAL: Demonstrate rendering the whole worksheet to PNG at 300 DPI
                string worksheetImage = "Worksheet_300dpi.png";
                ImageOrPrintOptions sheetImgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Png,
                    HorizontalResolution = 300,
                    VerticalResolution = 300
                };
                SheetRender sheetRender = new SheetRender(worksheet, sheetImgOptions);
                sheetRender.ToImage(0, worksheetImage);
                Console.WriteLine($"Worksheet page saved as \"{worksheetImage}\" with 300 DPI.");

            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during processing:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
