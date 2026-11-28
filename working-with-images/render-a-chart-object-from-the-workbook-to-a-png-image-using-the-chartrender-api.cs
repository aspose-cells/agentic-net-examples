// Title: How to export an Excel chart to a high‑resolution PNG image using Aspose.Cells ChartRender API in C#
// Common Searches: C# Aspose.Cells render chart as PNG with 300 DPI | Export Excel chart to PNG file using Aspose.Cells ChartRender | Save workbook chart as image in .NET with Aspose.Cells | Aspose.Cells Chart.ToImage example for high‑resolution PNG | Generate PNG from Excel chart programmatically in C#
// Tags: Aspose.Cells chart rendering to PNG | C# export Excel chart image | ChartRender API high resolution PNG | ImageOrPrintOptions PNG settings | ToImage method Aspose.Cells chart

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// The sample loads an existing workbook or creates one with a column chart, configures ImageOrPrintOptions for 300 dpi PNG output, and uses the Chart.ToImage method from Aspose.Cells ChartRender API to render the first worksheet chart to a file named "chart.png", with error handling for loading and rendering steps.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the workbook that contains the chart.
            // If the file does not exist, a sample workbook with a chart will be created.
            string workbookPath = "ChartWorkbook.xlsx";
            Workbook workbook;

            if (File.Exists(workbookPath))
            {
                workbook = new Workbook(workbookPath);
                Console.WriteLine($"Loaded workbook from \"{workbookPath}\".");
            }
            else
            {
                // Create a new workbook and add sample data.
                workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart.
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";
                chart.Title.Text = "Sample Chart";

                // Save the sample workbook for future runs.
                workbook.Save(workbookPath);
                Console.WriteLine($"Created sample workbook and saved to \"{workbookPath}\".");
            }

            // Retrieve the first chart from the first worksheet.
            Worksheet firstSheet = workbook.Worksheets[0];
            if (firstSheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the workbook. Exiting.");
                return;
            }

            Chart chartToRender = firstSheet.Charts[0];

            // Configure image options for PNG output.
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png,
                HorizontalResolution = 300,
                VerticalResolution = 300
            };

            // Render the chart to a PNG file.
            try
            {
                string outputImagePath = "chart.png";

                // Use Chart's ToImage method (available in recent Aspose.Cells versions).
                chartToRender.ToImage(outputImagePath, imgOptions);

                Console.WriteLine($"Chart successfully rendered to \"{outputImagePath}\".");
            }
            catch (Exception renderEx)
            {
                Console.WriteLine("Error during chart rendering:");
                Console.WriteLine(renderEx.Message);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred:");
            Console.WriteLine(ex.Message);
            Console.WriteLine(ex.StackTrace);
        }
    }
}
