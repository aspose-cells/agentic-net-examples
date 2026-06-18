using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class PlotAreaPictureFillDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

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

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // ------------------------------------------------------------
                // Apply picture fill to the plot area (not the whole chart)
                // ------------------------------------------------------------

                // Access the plot area's Area object
                Area plotAreaArea = chart.PlotArea.Area;

                // Set the fill type to Texture (used for picture fills)
                plotAreaArea.FillFormat.FillType = FillType.Texture;

                // Choose a picture fill mode (e.g., Stretch)
                plotAreaArea.FillFormat.PictureFormatType = FillPictureType.Stretch;

                // Provide image data – a tiny white pixel encoded in Base64
                string base64Image = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=";
                byte[] imageBytes = Convert.FromBase64String(base64Image);
                plotAreaArea.FillFormat.ImageData = imageBytes;

                // Optional: adjust transparency (0.0 = opaque, 1.0 = fully transparent)
                plotAreaArea.FillFormat.Transparency = 0.0;

                // Save the workbook to a file
                string outputPath = "PlotAreaPictureFillDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Run error: {ex.Message}");
                throw;
            }
        }
    }
}