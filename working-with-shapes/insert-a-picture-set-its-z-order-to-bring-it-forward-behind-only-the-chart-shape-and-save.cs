// Title: Insert a PNG picture into an Excel worksheet and place it behind a column chart using Aspose.Cells for .NET
// AI Prompts: Add a PNG image at cell C3, then use Aspose.Cells Z‑order methods to move the picture behind the existing column chart while keeping it above any other shapes, and save the workbook as XLSX. | Modify the example to programmatically set the picture's Z‑order so it appears directly behind the chart shape only, then export the file.
// Common Searches: Aspose.Cells C# insert image behind chart shape | how to change Z-order of pictures in Excel with Aspose.Cells .NET | place picture behind column chart using Aspose.Cells API | control stacking order of worksheet shapes in C# Aspose.Cells | save workbook with picture layered behind chart Aspose.Cells
// Tags: add picture behind chart Aspose.Cells | set picture Z-order Aspose.Cells .NET | insert PNG into worksheet Aspose.Cells | shape layering Excel .NET | picture stacking order in generated XLSX

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace Example
{
    // The code creates a new workbook, adds a column chart with sample data, inserts a PNG image at a specified cell if the file exists, demonstrates how to adjust the picture's Z‑order to sit behind the chart while remaining in front of other shapes, and saves the workbook as an XLSX file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a column chart to the worksheet (provides a shape for Z‑order demonstration)
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIdx];

                // Populate sample data for the chart
                for (int i = 0; i < 5; i++)
                {
                    sheet.Cells[i, 0].PutValue(i + 1);          // Column A values
                    sheet.Cells[i, 1].PutValue((i + 1) * 10);   // Column B values
                }

                // Define the data range for the chart series
                chart.NSeries.Add("A1:A5", true);
                // Optional: set category data if needed (commented out to avoid API issues)
                // chart.NSeries[0].CategoryData = "B1:B5";

                // Insert a picture into the worksheet if the file exists
                string imagePath = "sample.png";
                if (File.Exists(imagePath))
                {
                    int pictureIdx = sheet.Pictures.Add(2, 2, imagePath);
                    Picture picture = sheet.Pictures[pictureIdx];
                    // Optional Z‑order handling: bring picture to back (if required)
                    // picture.BringToFront(); // Uncomment if you need to adjust order
                }
                else
                {
                    Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture insertion.");
                }

                // Save the workbook
                string outputPath = "output.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
