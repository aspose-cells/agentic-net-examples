// Title: C# – Insert Picture into Excel Chart and Adjust Z‑Order with Aspose.Cells
// Description: Shows how to create a workbook, add sample data and a column chart, load a JPEG file, insert the image into the chart at defined coordinates, move the picture one step forward in the Z‑order so it sits just behind the chart shape, and save the result as an XLSX file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# | insert picture into Excel chart | chart image layering | Z order shape Aspose.Cells | ToFrontOrBack method | AddPictureInChart | Excel chart picture .NET | Aspose.Cells example | C# Excel chart image | Aspose.Cells Z‑order tutorial | US developers | India developers | GitHub Aspose.Cells sample
// Common Searches: How to add an image inside a chart with Aspose.Cells for .NET | Aspose.Cells set Z‑order of picture in Excel chart | Insert picture in chart and control layering Aspose.Cells | Bring picture forward one level behind chart shape Aspose.Cells | C# code to use ToFrontOrBack with chart shapes | Aspose.Cells example on GitHub for picture Z‑order
// Developer Intent: Add an image to a chart and set its Z‑order so it appears directly behind the chart series.
// Use Cases: Place a company logo on a chart while keeping data series on top. | Add a subtle watermark to a chart without obscuring data points. | Create a branded chart with a background picture that sits behind plotted values.
// AI Prompts: Generate C# code using Aspose.Cells that loads a JPEG, inserts it into a chart at specific coordinates, and moves the picture forward by one Z‑order step. | Explain how the ToFrontOrBack method controls shape layering when adding a picture to an Excel chart in Aspose.Cells. | Provide a complete Aspose.Cells example that inserts an image into a chart, sets its size, and ensures the chart data remains on top of the picture.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Shows how to create a workbook, add sample data and a column chart, load a JPEG file, insert the image into the chart at defined coordinates, move the picture one step forward in the Z‑order so it sits just behind the chart shape, and save the result as an XLSX file using Aspose.Cells for .NET.
class InsertPictureWithZOrder
{
    static void Main()
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

            // Add a column chart to the worksheet
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 1, 20, 10);
            Chart chart = sheet.Charts[chartIdx];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Path to the picture file
            string imagePath = "example.jpg";

            // Verify that the image file exists before loading
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Load picture data and insert it into the chart
            byte[] imgBytes = File.ReadAllBytes(imagePath);
            using (MemoryStream imgStream = new MemoryStream(imgBytes))
            {
                // Add picture to the chart (position and size are in points)
                Picture pic = chart.Shapes.AddPictureInChart(100, 100, imgStream, 50, 50);

                // Bring the picture forward by one position in the Z‑order
                pic.ToFrontOrBack(1);
            }

            // Save the workbook
            workbook.Save("PictureWithZOrder.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
