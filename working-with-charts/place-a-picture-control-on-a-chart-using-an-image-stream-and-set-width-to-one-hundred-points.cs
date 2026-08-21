// Title: C# – Add a picture to an Aspose.Cells chart from a MemoryStream and set its width to 100 pt
// Description: Creates a new workbook, builds a column chart with sample data, loads an image into a MemoryStream, inserts the image into the chart using AddPictureInChart, sets the picture width to 100 points, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells C# chart image | AddPictureInChart MemoryStream | set picture width points Aspose.Cells | chart picture control .NET | insert image into Excel chart programmatically | Aspose.Cells example | Excel chart branding logo
// Common Searches: Aspose.Cells add picture to chart from stream | C# set picture width in points on chart | How to use AddPictureInChart in Aspose.Cells | Insert image into Excel chart using Aspose.Cells .NET | Aspose.Cells chart picture control example
// Developer Intent: Insert an image into a chart from a MemoryStream and define its width as 100 points using Aspose.Cells for .NET.
// Use Cases: Brand a generated chart with a company logo for consistent visual identity. | Add a dynamic watermark or badge to charts based on runtime data. | Place custom icons next to specific data points in automated reports.
// AI Prompts: Show how to change the picture height and position after adding it to an Aspose.Cells chart. | Provide code to download an image from a URL, load it into a MemoryStream, and add it to a chart with custom offsets. | Explain the scaling parameters of AddPictureInChart and how to retrieve the Picture object for further formatting.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a new workbook, builds a column chart with sample data, loads an image into a MemoryStream, inserts the image into the chart using AddPictureInChart, sets the picture width to 100 points, and saves the file as an Excel workbook.
class Program
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
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Load an image file into a memory stream if it exists
            const string imagePath = "example.jpg";
            if (File.Exists(imagePath))
            {
                byte[] imageData = File.ReadAllBytes(imagePath);
                using (MemoryStream imgStream = new MemoryStream(imageData))
                {
                    // Add the picture to the chart using the stream.
                    // Offsets are in 1/4000 of the chart area; width/height scales are percentages.
                    Picture picture = chart.Shapes.AddPictureInChart(0, 0, imgStream, 100, 100);
                    picture.WidthPt = 100; // Set picture width to 100 points
                }
            }
            else
            {
                Console.WriteLine($"Warning: Image file '{imagePath}' not found. Skipping picture insertion.");
            }

            // Save the workbook with the chart (and picture if added)
            workbook.Save("ChartWithPicture.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
