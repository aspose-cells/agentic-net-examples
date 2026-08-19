// Title: Make a Chart 40% Transparent Over a Worksheet Background Image with Aspose.Cells for .NET
// Description: This example shows how to add a PNG background picture to a worksheet, create a column chart, and set the chart area to 40% transparency using Aspose.Cells for .NET before saving the workbook.
// Keywords: Aspose.Cells chart transparency | C# chart opacity 40 percent | overlay chart on worksheet image | BackgroundMode Transparent Aspose.Cells | Area.Transparency property .NET | semi‑transparent chart Excel | Aspose.Cells chart styling
// Common Searches: Aspose.Cells set chart transparency | C# make chart semi transparent in Excel | how to overlay chart on worksheet background | chart area opacity 40% Aspose.Cells | transparent chart area Aspose.Cells .NET
// Developer Intent: Apply 40% transparency to a chart so it blends with a worksheet background image.
// Use Cases: Design a sales dashboard where the chart subtly reveals a watermark behind it. | Create a financial report that places a semi‑transparent chart over a branded background graphic. | Produce a presentation‑style Excel file with charts that appear partially see‑through for visual depth.
// AI Prompts: Generate C# code using Aspose.Cells to insert a background picture and render a column chart with 40% transparency. | Explain the effect of BackgroundMode = Transparent and Area.Transparency on chart rendering in Aspose.Cells. | Provide step‑by‑step instructions to adjust a chart's opacity to a specific percentage and save the workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartTransparencyDemo
{
    // This example shows how to add a PNG background picture to a worksheet, create a column chart, and set the chart area to 40% transparency using Aspose.Cells for .NET before saving the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a background image to the worksheet if the file exists
                string backgroundPath = "background.png";
                if (File.Exists(backgroundPath))
                {
                    int pictureIndex = worksheet.Pictures.Add(0, 0, backgroundPath);
                    // Optional: adjust picture placement
                    // worksheet.Pictures[pictureIndex].Placement = PlacementType.FreeFloating;
                }
                else
                {
                    Console.WriteLine($"Warning: Background image '{backgroundPath}' not found. Skipping picture insertion.");
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

                // Create a chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Set chart area transparency to 40%
                chart.ChartArea.BackgroundMode = BackgroundMode.Transparent;
                chart.ChartArea.Area.Transparency = 0.4; // 0.0 = opaque, 1.0 = fully clear

                // Save the workbook
                string outputPath = "ChartWithTransparency.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
