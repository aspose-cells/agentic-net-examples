// Title: Insert a picture control into a column chart from a memory stream and set its width to 100 points using Aspose.Cells for .NET (C#)
// Common Searches: how to add a picture control to an Excel chart using a byte array with Aspose.Cells in C# | setting picture width in points for a chart picture control in Aspose.Cells | using MemoryStream to insert an image into a column chart with Aspose.Cells .NET | C# code to place a picture over a chart area in an Aspose.Cells workbook | Aspose.Cells example for adding and resizing a chart picture control
// Tags: picture control from memory stream Aspose.Cells | chart picture width points Aspose.Cells | insert image into column chart C# Aspose.Cells | resize chart picture programmatically Aspose.Cells | worksheet picture placement Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The sample loads an image (from a file or a fallback PNG), creates a workbook, adds sample data, builds a column chart, inserts the image as a picture control on the worksheet using a MemoryStream, sets the picture's width to 100 points, and saves the workbook as an Excel file.
class Program
{
    static void Main()
    {
        try
        {
            // Path to an image file that will be used for the picture control.
            // Replace with a valid image path on your machine or ensure the file exists.
            string imagePath = "sample.png";

            // Load image bytes – if the file does not exist we use a minimal 1x1 red PNG.
            byte[] imageBytes;
            if (File.Exists(imagePath))
            {
                imageBytes = File.ReadAllBytes(imagePath);
            }
            else
            {
                // Base64‑encoded 1x1 red PNG.
                const string redPngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wH+X6ZLAAAAAElFTkSuQmCC";
                imageBytes = Convert.FromBase64String(redPngBase64);
            }

            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data for the chart.
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["A4"].PutValue("Cherry");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(45);
            sheet.Cells["B4"].PutValue(25);

            // Add a column chart.
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Add a picture to the worksheet (as a picture control) using an image stream.
            // Position it roughly over the chart area.
            int pictureIndex;
            try
            {
                pictureIndex = sheet.Pictures.Add(5, 0, new MemoryStream(imageBytes));
            }
            catch (Exception picEx)
            {
                Console.WriteLine($"Failed to add picture: {picEx.Message}");
                return;
            }

            Picture picture = sheet.Pictures[pictureIndex];
            picture.Width = 100; // Set width to 100 points.

            // Save the workbook.
            string outputPath = "ChartWithPictureControl.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            Console.WriteLine("Picture control added (as a worksheet picture) with width set to 100 points.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
