// Title: Aspose.Cells for .NET – Set Chart Transparency to 40 % for Overlay on Worksheet Image
// Description: Shows how to create a workbook, optionally add a background picture, insert sample data, generate a column chart, set the chart background to Transparent, apply 40 % area transparency, and save the result. Perfect for layering charts over maps, logos, or photos.
// Keywords: Aspose.Cells | .NET | chart transparency | chart opacity | ChartArea.Transparency | transparent chart background | overlay chart on image | Excel chart overlay | PlacementType.FreeFloating | background picture Aspose.Cells | Excel dashboard transparency
// Common Searches: Aspose.Cells set chart opacity to 40 percent | How to make a chart semi‑transparent in .NET | Overlay Excel chart on background image using Aspose.Cells | Transparent chart background Aspose.Cells .NET | ChartArea.Transparency example Aspose.Cells
// Developer Intent: Apply a 40 % transparency to a chart so it can be placed over a worksheet image without obscuring the picture.
// Use Cases: Create a sales dashboard where a column chart blends with a company logo placed behind it. | Layer a performance chart on a geographic map to reveal underlying regions. | Design an Excel slide that combines a chart with a background photograph for a polished presentation.
// AI Prompts: Generate C# code to set a chart’s transparency to 30 % with Aspose.Cells. | Provide an example of overlaying a pie chart on a worksheet picture and making the chart background transparent. | Explain the effect of PlacementType.FreeFloating on the stacking order of charts and pictures in Aspose.Cells.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartTransparencyDemo
{
    // Shows how to create a workbook, optionally add a background picture, insert sample data, generate a column chart, set the chart background to Transparent, apply 40 % area transparency, and save the result. Perfect for layering charts over maps, logos, or photos.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Path to the background image
                string backgroundImagePath = "background.png";

                // Insert a background picture if the file exists
                if (File.Exists(backgroundImagePath))
                {
                    int pictureIndex = worksheet.Pictures.Add(0, 0, backgroundImagePath);
                    // Set picture to free floating so it stays behind the chart
                    worksheet.Pictures[pictureIndex].Placement = PlacementType.FreeFloating;
                }
                else
                {
                    Console.WriteLine($"Warning: Background image '{backgroundImagePath}' not found. Skipping picture insertion.");
                }

                // Add sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);

                // Add a column chart that will be placed over the picture
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Make the chart background transparent and set its area transparency to 40%
                chart.ChartArea.BackgroundMode = BackgroundMode.Transparent;
                chart.ChartArea.Area.Transparency = 0.4; // 0.0 = opaque, 1.0 = clear

                // Save the workbook
                string outputPath = "ChartWithTransparency.xlsx";
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
