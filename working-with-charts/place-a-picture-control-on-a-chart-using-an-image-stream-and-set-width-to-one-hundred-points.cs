// Title: Add a picture to an Aspose.Cells chart from a MemoryStream and set width to 100 pt (C#)
// Description: Creates a new workbook, builds a column chart with sample data, loads an image file into a MemoryStream, inserts the image onto the chart using AddPictureInChart, sets the picture's WidthPt property to 100 points, and saves the workbook as an Excel file.
// Keywords: Aspose.Cells | C# | AddPictureInChart | chart picture insertion | MemoryStream image | WidthPt | set picture width points | Excel chart overlay | image stream to chart | Aspose.Cells chart customization
// Common Searches: How to add an image to a chart with Aspose.Cells C# | Set picture width in points on an Aspose.Cells chart | Insert picture from MemoryStream into Excel chart using Aspose | Aspose.Cells AddPictureInChart example code | Resize chart picture to exact points in Aspose.Cells
// Developer Intent: Insert an image into a chart from a memory stream and define its width as exactly 100 points.
// Use Cases: Brand a generated chart with a company logo for corporate reports. | Apply a confidential watermark image over a chart before distribution. | Show product thumbnails next to data points in a sales performance chart.
// AI Prompts: Provide C# code that loads an image into a MemoryStream and adds it to an Aspose.Cells chart with a width of 100 pt. | Explain the offset parameters of AddPictureInChart and how to calculate them for precise placement. | Show robust error‑handling for missing image files when inserting pictures into Aspose.Cells charts.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a new workbook, builds a column chart with sample data, loads an image file into a MemoryStream, inserts the image onto the chart using AddPictureInChart, sets the picture's WidthPt property to 100 points, and saves the workbook as an Excel file.
class AddPictureToChart
{
    static void Main()
    {
        try
        {
            // Create a new workbook and obtain the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Load an image file into a memory stream if it exists
            const string imagePath = "example.png";
            if (File.Exists(imagePath))
            {
                try
                {
                    byte[] imageBytes = File.ReadAllBytes(imagePath);
                    using (MemoryStream imageStream = new MemoryStream(imageBytes))
                    {
                        // Add the picture to the chart (offsets are in 1/4000 of the chart area)
                        // widthScale and heightScale are set to 50% as an initial size
                        Picture picture = chart.Shapes.AddPictureInChart(100, 100, imageStream, 50, 50);
                        // Set the picture width to exactly 100 points
                        picture.WidthPt = 100;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error loading image '{imagePath}': {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture insertion.");
            }

            // Save the workbook with the chart (and picture if added)
            workbook.Save("ChartWithPicture.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
