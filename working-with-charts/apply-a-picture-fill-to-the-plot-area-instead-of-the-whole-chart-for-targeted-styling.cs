// Title: How to apply a picture texture fill to a chart's plot area with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set image as plot area background in Excel chart | add picture fill to chart plot area using Aspose.Cells .NET | C# example of texture fill for chart plot area with Aspose.Cells | how to use FillFormat.ImageData for plot area in Aspose.Cells | apply custom image to Excel chart plot area programmatically
// Tags: plotarea texture fill Aspose.Cells | chart plot area image background C# | FillFormat FillType Texture usage | Aspose.Cells set plotarea picture fill | Excel chart plot area styling .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace PlotAreaPictureFillDemo
{
    // The sample loads a PNG file into a byte array, creates a workbook with sample data, adds a column chart, and applies the image as a texture fill to the chart's PlotArea by setting FillFormat.FillType to Texture and assigning the image data, then saves the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the picture that will be used as fill for the plot area.
            // Change this path to point to an existing image file on your machine.
            string picturePath = "plotarea_fill.png";

            if (!File.Exists(picturePath))
            {
                Console.WriteLine($"Error: Picture file not found at '{picturePath}'.");
                return;
            }

            // Load the picture into a byte array.
            byte[] pictureData;
            try
            {
                pictureData = File.ReadAllBytes(picturePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading picture file: {ex.Message}");
                return;
            }

            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data for the chart.
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

            // Set the data range for the chart.
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // ------------------------------------------------------------
            // Apply picture fill to the PlotArea only (not the whole chart).
            // ------------------------------------------------------------

            // Access the PlotArea of the chart.
            PlotArea plotArea = chart.PlotArea;

            // Set the fill type to Texture (used for image fills in this version).
            plotArea.Area.FillFormat.FillType = FillType.Texture;

            // Assign the picture data.
            plotArea.Area.FillFormat.ImageData = pictureData;

            // Optional: adjust transparency or scaling if desired.
            // plotArea.Area.FillFormat.Transparency = 0.2; // 20% transparent
            // plotArea.Area.FillFormat.Scale = 1.0;       // 100% size

            // Save the workbook.
            string outputPath = "ChartWithPlotAreaPictureFill.xlsx";
            try
            {
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
    }
}
