// Title: Make an Excel chart 40% transparent and overlay it on a worksheet background image using Aspose.Cells for .NET
// AI Prompts: Write C# code that adds a PNG picture as a free‑floating background to a worksheet and sets the chart’s opacity to 40 % using Aspose.Cells. | Show how to configure the ChartArea and PlotArea objects to be 40 % transparent while positioning the chart over an inserted image in an Excel workbook with Aspose.Cells for .NET. | Generate a complete Aspose.Cells example that creates a column chart, binds sample data, inserts a background picture, and applies 40 % transparency to both the chart and plot areas.
// Common Searches: asp.net aspose.cells make chart semi transparent over background picture | c# set chart opacity 40 percent in Excel using Aspose.Cells | how to place a chart on top of an image in an Excel file with Aspose.Cells | adjust plot area opacity in Aspose.Cells chart example | Aspose.Cells overlay chart on worksheet image tutorial
// Tags: chart opacity setting Aspose.Cells | plot area opacity Aspose.Cells | free floating picture insertion Aspose.Cells | column chart data binding Aspose.Cells C# | excel workbook save with transparent chart Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsChartTransparencyDemo
{
    // Creates a new workbook, optionally inserts a PNG as a free‑floating background picture, adds a column chart with sample data, sets both the chart area and plot area opacity to 40 %, and saves the file as ChartWithTransparency.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Insert an image that will serve as the background if the file exists
                const string backgroundPath = "background.png";
                if (File.Exists(backgroundPath))
                {
                    int pictureIndex = worksheet.Pictures.Add(0, 0, backgroundPath);
                    // Set picture to free floating so it can be positioned independently
                    worksheet.Pictures[pictureIndex].Placement = PlacementType.FreeFloating;
                }
                else
                {
                    Console.WriteLine($"Warning: Background image '{backgroundPath}' not found. Skipping image insertion.");
                }

                // Add a sample chart that will be placed over the image
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);

                // Bind data to the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Make the chart area transparent (40% transparent)
                ChartArea chartArea = chart.ChartArea;
                chartArea.BackgroundMode = BackgroundMode.Transparent;
                chartArea.Area.Transparency = 0.4; // 0.0 = opaque, 1.0 = fully clear

                // Also make the plot area transparent for a cleaner overlay
                PlotArea plotArea = chart.PlotArea;
                plotArea.Area.Transparency = 0.4;

                // Save the workbook
                const string outputPath = "ChartWithTransparency.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
