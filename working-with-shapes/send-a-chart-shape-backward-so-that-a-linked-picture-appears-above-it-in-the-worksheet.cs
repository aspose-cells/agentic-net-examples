// Title: Send a ChartShape to the back so a linked picture overlays it – Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, add a column chart, insert a linked PNG image, and call ChartShape.ToFrontOrBack(0) to move the chart backward in the Z‑order so the picture appears on top before saving the file.
// Keywords: Aspose.Cells ChartShape Z-order | C# chart backward Aspose.Cells | overlay picture on chart Excel | ChartShape ToFrontOrBack example | Aspose.Cells shape layering | Excel chart picture stacking | Aspose.Cells .NET image over chart
// Common Searches: Aspose.Cells move chart to back | How to overlay image on chart using Aspose.Cells C# | ChartShape Z order Aspose.Cells | Place logo over chart Aspose.Cells | Send chart shape backward Aspose.Cells
// Developer Intent: Move the chart shape behind a linked picture so the picture is displayed on top of the chart.
// Use Cases: Add a watermark that must sit above chart graphics in generated reports. | Place a company logo over a sales chart to reinforce branding. | Create Excel dashboards where annotations or icons need to overlay existing charts.
// AI Prompts: Write C# code with Aspose.Cells that adds a chart, inserts a picture, and uses ChartShape.ToFrontOrBack to place the picture above the chart. | Explain how the ToFrontOrBack method controls Z‑order for shapes in Aspose.Cells and give examples for moving shapes forward and backward. | Provide a snippet that iterates over all shapes in a worksheet and sets their Z‑order so specific images appear on top of charts.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartZOrderExample
{
    // Shows how to create a workbook, add a column chart, insert a linked PNG image, and call ChartShape.ToFrontOrBack(0) to move the chart backward in the Z‑order so the picture appears on top before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // Add a chart to the worksheet
                // -------------------------------------------------
                // Add a column chart spanning rows 5-15 and columns 0-5
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = worksheet.Charts[chartIndex];

                // Add a simple data series (replace with actual data range as needed)
                chart.NSeries.Add("A1:A5", true);

                // Get the ChartShape (the visual representation of the chart)
                ChartShape chartShape = chart.ChartObject;

                // -------------------------------------------------
                // Add a picture (linked image) to the worksheet
                // -------------------------------------------------
                string imagePath = "linkedImage.png";
                if (File.Exists(imagePath))
                {
                    int pictureIndex = worksheet.Pictures.Add(0, 0, imagePath);
                    Picture picture = worksheet.Pictures[pictureIndex];
                }
                else
                {
                    Console.WriteLine($"Image file not found: {imagePath}. Skipping picture insertion.");
                }

                // -------------------------------------------------
                // Send the chart shape backward so the picture appears on top
                // -------------------------------------------------
                // 0 moves the shape toward the back of the Z-order
                chartShape.ToFrontOrBack(0);

                // -------------------------------------------------
                // Save the workbook
                // -------------------------------------------------
                string outputPath = "ChartWithPictureOnTop.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
