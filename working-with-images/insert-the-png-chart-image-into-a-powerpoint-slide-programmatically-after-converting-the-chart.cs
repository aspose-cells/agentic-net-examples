// Title: Convert an Aspose.Cells column chart to a high‑resolution PNG image for PowerPoint slide insertion using C#
// Common Searches: how to export an Aspose.Cells chart as a PNG for PowerPoint in C# | Aspose.Cells render chart to high resolution PNG using memory stream | C# code to save an Excel column chart as a 300 DPI PNG image | convert Excel chart to PNG for slide insertion with Aspose.Cells | using ImageOrPrintOptions to generate PNG from an Aspose.Cells chart
// Tags: Aspose.Cells chart PNG export for PowerPoint | C# high‑resolution chart image rendering | ImageOrPrintOptions 300 DPI PNG | chart.ToImage memory stream usage | save chart image file C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace InsertChartToPowerPoint
{
    // The example creates a workbook, adds sample data and a column chart, then renders the chart to a PNG image with 300 dpi resolution using ImageOrPrintOptions. The PNG is written to a MemoryStream, saved as 'ChartImage.png', and the workbook is also saved, providing a high‑quality image ready for insertion into a PowerPoint slide.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a workbook and add sample data for the chart.
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Header
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");

                // Data rows
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(1200);
                sheet.Cells["A3"].PutValue("Orange");
                sheet.Cells["B3"].PutValue(800);
                sheet.Cells["A4"].PutValue("Banana");
                sheet.Cells["B4"].PutValue(1500);

                // -----------------------------------------------------------------
                // 2. Add a column chart and bind it to the data range.
                // -----------------------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";      // Categories
                chart.Title.Text = "Fruit Sales";

                // -----------------------------------------------------------------
                // 3. Convert the chart to a PNG image using a MemoryStream.
                // -----------------------------------------------------------------
                using (MemoryStream chartImageStream = new MemoryStream())
                {
                    ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                    {
                        ImageType = ImageType.Png,
                        HorizontalResolution = 300,
                        VerticalResolution = 300
                    };

                    // Render the chart to the memory stream.
                    chart.ToImage(chartImageStream, imgOptions);
                    chartImageStream.Position = 0; // Reset for reading.

                    // -----------------------------------------------------------------
                    // 4. Save the chart image to disk for verification.
                    // -----------------------------------------------------------------
                    string imagePath = "ChartImage.png";

                    // Ensure we don't overwrite an existing file unintentionally.
                    if (File.Exists(imagePath))
                    {
                        File.Delete(imagePath);
                    }

                    using (FileStream file = new FileStream(imagePath, FileMode.Create, FileAccess.Write))
                    {
                        chartImageStream.CopyTo(file);
                    }

                    Console.WriteLine($"Chart image saved successfully to '{imagePath}'.");
                }

                // -----------------------------------------------------------------
                // 5. (Optional) Save the workbook if further inspection is needed.
                // -----------------------------------------------------------------
                string workbookPath = "ChartWorkbook.xlsx";
                if (File.Exists(workbookPath))
                {
                    File.Delete(workbookPath);
                }
                workbook.Save(workbookPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{workbookPath}'.");
            }
            catch (Exception ex)
            {
                // Report any errors that occur during processing.
                Console.Error.WriteLine($"Error: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
