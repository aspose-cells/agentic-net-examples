// Title: Aspose.Cells .NET – Retrieve Chart X/Y Pixel Position and Size via ChartShape
// Description: Demonstrates how to obtain a chart's top‑left X and Y pixel coordinates, width, and height using the ChartShape (ChartObject) in Aspose.Cells for .NET, and shows an optional conversion to the internal 1/4000‑unit system for exact layout control before saving the workbook.
// Keywords: Aspose.Cells chart position | ChartShape X Y coordinates .NET | retrieve chart pixel location | chart size Aspose.Cells | convert chart coordinates 1/4000 units | C# Aspose.Cells chart alignment | Aspose.Cells ChartObject properties
// Common Searches: how to get chart X Y pixel coordinates Aspose.Cells | Aspose.Cells chart shape position and dimensions | convert chart pixel values to 1/4000 units C# | retrieve chart size for alignment Aspose.Cells | Aspose.Cells get chart top left corner
// Developer Intent: Extract a chart’s pixel‑based location and dimensions to enable precise programmatic alignment within a worksheet.
// Use Cases: Place another shape (image, textbox) relative to an existing chart by using its X, Y, width, and height. | Dynamically adjust worksheet layout when adding or moving charts based on exact pixel offsets. | Map pixel coordinates to Aspose.Cells’ 1/4000 unit system for accurate positioning of annotations or data points.
// AI Prompts: Show C# code that reads a chart’s X, Y, width, and height using Aspose.Cells and converts the values to 1/4000 units. | Provide an example of aligning a picture shape next to a chart by using the chart’s pixel coordinates in Aspose.Cells for .NET. | Explain how to use ChartShape properties to reposition a chart for pixel‑perfect layout in an Excel file with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to obtain a chart's top‑left X and Y pixel coordinates, width, and height using the ChartShape (ChartObject) in Aspose.Cells for .NET, and shows an optional conversion to the internal 1/4000‑unit system for exact layout control before saving the workbook.
    public class RetrieveChartCoordinates
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);

                // Add a column chart to the worksheet (rows 5‑20, columns 0‑8)
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // OPTIONAL: calculate the chart if you need point‑level coordinates later
                chart.Calculate();

                // Retrieve the chart shape (the container of the chart on the worksheet)
                ChartShape chartShape = chart.ChartObject;

                // X and Y coordinates of the upper‑left corner of the chart shape (pixels)
                int chartXPixel = chartShape.X; // horizontal offset from worksheet left border
                int chartYPixel = chartShape.Y; // vertical offset from worksheet top border

                // Width and height of the chart shape (pixels) – useful for further alignment calculations
                int chartWidthPixel = chartShape.Width;
                int chartHeightPixel = chartShape.Height;

                // Display the retrieved coordinates
                Console.WriteLine($"Chart upper‑left corner: X = {chartXPixel} px, Y = {chartYPixel} px");
                Console.WriteLine($"Chart size: Width = {chartWidthPixel} px, Height = {chartHeightPixel} px");

                // Example: convert the pixel coordinates to the internal 1/4000 unit system
                // (used by properties like ShapeX, ShapeY of ChartPoint)
                double xIn4000Units = chartXPixel * 4000.0 / chartWidthPixel;
                double yIn4000Units = chartYPixel * 4000.0 / chartHeightPixel;
                Console.WriteLine($"Chart upper‑left corner in 1/4000 units: X = {xIn4000Units:F2}, Y = {yIn4000Units:F2}");

                // Save the workbook (ensure the directory exists)
                string outputPath = "ChartCoordinatesDemo.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            RetrieveChartCoordinates.Run();
        }
    }
}
