// Title: Add a picture to an Aspose.Cells chart from a MemoryStream and set its width to 100 pt (C#)
// Description: This example creates a workbook, populates it with sample data, generates a column chart, loads an image file into a MemoryStream, inserts the image onto the chart with AddPictureInChart, forces the picture width to 100 points (height scales automatically), and saves the file as ChartWithPicture.xlsx.
// Keywords: Aspose.Cells chart picture | AddPictureInChart C# | MemoryStream image Aspose.Cells | set picture width points | .NET chart shape width | Aspose.Cells image overlay | C# Excel chart logo | Aspose.Cells US developers | Aspose.Cells Europe examples
// Common Searches: how to insert an image into an Aspose.Cells chart using C# | set picture width in points on an Aspose.Cells chart | add picture to chart from a stream Aspose.Cells | Aspose.Cells AddPictureInChart width 100 pt | C# place logo on Excel chart with Aspose
// Developer Intent: Insert an image onto a chart from a stream and define its width in points.
// Use Cases: Brand a sales chart with a company logo for presentations. | Show a product thumbnail next to a performance chart generated at runtime. | Add a confidentiality watermark to financial charts before distribution.
// AI Prompts: Write C# code that reads an image byte array, creates a MemoryStream, and adds the picture to an Aspose.Cells chart with a width of 100 points. | Explain how Aspose.Cells automatically scales picture height when WidthPt is set on a chart shape. | Provide error‑handling patterns for missing image files when using AddPictureInChart with a MemoryStream.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// This example creates a workbook, populates it with sample data, generates a column chart, loads an image file into a MemoryStream, inserts the image onto the chart with AddPictureInChart, forces the picture width to 100 points (height scales automatically), and saves the file as ChartWithPicture.xlsx.
class AddPictureToChart
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 1, 20, 10);
            Chart chart = sheet.Charts[chartIdx];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Path to the image file
            string imagePath = "example.jpg";

            // Load image data into a memory stream if the file exists
            if (File.Exists(imagePath))
            {
                byte[] imgBytes = File.ReadAllBytes(imagePath);
                using (MemoryStream imgStream = new MemoryStream(imgBytes))
                {
                    // Add picture to the chart; offsets are in 1/4000 of chart area
                    // WidthScale and HeightScale are percentages (100 = original size)
                    Picture pic = chart.Shapes.AddPictureInChart(100, 100, imgStream, 100, 100);
                    // Set the picture width to 100 points (height scales proportionally)
                    pic.WidthPt = 100;
                }
            }
            else
            {
                Console.WriteLine($"Image file not found: {imagePath}. Skipping picture insertion.");
            }

            // Save the workbook
            workbook.Save("ChartWithPicture.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
