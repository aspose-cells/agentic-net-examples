using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;   // Needed for FillType and FillPictureType enums

namespace AsposeCellsExamples
{
    public class PlotAreaPictureFillDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // ------------------------------------------------------------
                // Apply picture fill to the plot area (not the whole chart)
                // ------------------------------------------------------------

                // Access the plot area of the chart
                PlotArea plotArea = chart.PlotArea;

                // Set the fill type to Texture (required for picture fill)
                plotArea.Area.FillFormat.FillType = FillType.Texture;

                // Choose how the picture is applied (e.g., StackAndScale)
                plotArea.Area.FillFormat.PictureFormatType = FillPictureType.StackAndScale;

                // Load an image (for demo we use a small in‑memory PNG)
                // Here we create a 1x1 white pixel PNG from a base64 string.
                string base64Image = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=";
                byte[] imageData = Convert.FromBase64String(base64Image);
                plotArea.Area.FillFormat.ImageData = imageData;

                // ------------------------------------------------------------
                // Save the workbook
                // ------------------------------------------------------------
                string outputPath = "PlotAreaPictureFillDemo.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            PlotAreaPictureFillDemo.Run();
        }
    }
}