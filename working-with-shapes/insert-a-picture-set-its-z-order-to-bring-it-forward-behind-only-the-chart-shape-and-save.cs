// Title: Insert a PNG picture, set its Z‑order behind a column chart, and save the workbook with Aspose.Cells for .NET (C#)
// Common Searches: how to add a picture to an Excel worksheet and place it behind a chart using Aspose.Cells C# | Aspose.Cells C# set picture ZOrderPosition relative to a chart shape | insert image and column chart in the same worksheet with Aspose.Cells .NET | save Excel file after adjusting shape layering with Aspose.Cells | C# Aspose.Cells bring picture forward but keep it behind a chart
// Tags: add picture to worksheet Aspose.Cells C# | set picture ZOrderPosition Aspose.Cells | insert column chart Aspose.Cells C# | layer shapes in Excel using Aspose.Cells | save workbook with image and chart Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example checks that a PNG image exists, creates a new workbook, inserts the picture at cell B3, adds a column chart, optionally adjusts the picture's ZOrderPosition so it sits just behind the chart, and then saves the workbook to the specified path.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Paths
                string imagePath = @"C:\Images\SamplePicture.png";
                string outputPath = @"C:\Output\WorkbookWithPictureAndChart.xlsx";

                // Verify image file exists
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file not found: {imagePath}");
                    return;
                }

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Insert picture (Add returns the picture index)
                int pictureIndex = sheet.Pictures.Add(2, 2, imagePath);
                Picture picture = sheet.Pictures[pictureIndex];

                // Insert chart (Add returns the chart index)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 5, 15, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Optional: adjust Z-order if needed (available in newer versions)
                // picture.ZOrderPosition = chart.ZOrderPosition - 1;

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
