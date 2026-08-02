// Title: Send a chart shape backward so a picture appears on top with Aspose.Cells for .NET
// Description: Shows how to insert a column chart and an image into a worksheet, then push the chart behind the image using ChartShape.ToFrontOrBack, and finally save the workbook.
// Keywords: Aspose.Cells chart Z-order | ChartShape ToFrontOrBack | C# move chart behind image | Aspose.Cells picture over chart | Excel shape layering .NET
// Common Searches: Aspose.Cells move chart behind picture C# | Change Z-order of chart shape in Aspose.Cells | Place image over chart Aspose.Cells .NET | ChartShape ToFrontOrBack example | Layer shapes in Excel using Aspose.Cells
// Developer Intent: Place an image above a chart by sending the chart shape to the back.
// Use Cases: Add a company logo that overlays charts in financial reports. | Show annotation graphics above data visualizations in dashboards. | Apply a watermark image that covers chart areas in generated worksheets.
// AI Prompts: Provide C# code that uses Aspose.Cells to move a chart behind an inserted picture. | Explain how ChartShape.ToFrontOrBack controls Z-order of shapes in an Excel file. | Show an example of layering a picture over a chart with Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartZOrderDemo
{
    // Shows how to insert a column chart and an image into a worksheet, then push the chart behind the image using ChartShape.ToFrontOrBack, and finally save the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a column chart to the worksheet (rows 5-15, columns 0-5)
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = worksheet.Charts[chartIndex];
                ChartShape chartShape = chart.ChartObject;

                // Path to the image file to be added as a picture
                string imagePath = "image.png";

                // Verify that the image file exists before adding it
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file not found: {imagePath}");
                    return;
                }

                // Add the picture to the worksheet and obtain the Picture object
                int pictureIndex = worksheet.Pictures.Add(0, 0, imagePath);
                Picture picture = worksheet.Pictures[pictureIndex];

                // Send the chart shape backward so the picture appears on top
                chartShape.ToFrontOrBack(-1);

                // Save the workbook
                string outputPath = "ChartWithPictureZOrder.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
