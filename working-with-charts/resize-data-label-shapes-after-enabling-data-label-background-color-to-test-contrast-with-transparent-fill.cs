// Title: Resize data label shapes and set transparent background in a column chart using Aspose.Cells for .NET
// AI Prompts: Create a column chart, enable data labels, set BackgroundMode to Transparent, disable automatic shape resizing, and assign WidthPixel = 80 and HeightPixel = 30 with Aspose.Cells for .NET. | Modify an existing chart to turn off auto‑sizing of data label shapes, apply custom pixel dimensions, and keep a transparent fill using Aspose.Cells. | Change the font color and size of chart data labels that have a transparent background in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells set data label width and height in pixels | how to make chart data label background transparent in Aspose.Cells .NET | disable automatic resizing of data label shapes Aspose.Cells chart | customize font color of data labels with transparent background Aspose.Cells | column chart data label size adjustment Aspose.Cells example
// Tags: set data label shape size Aspose.Cells | transparent background for chart data labels .NET | disable data label auto resize Aspose.Cells | customize data label font color Aspose.Cells | column chart data label customization Aspose.Cells | pixel dimensions for data label shapes Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates creating a column chart, enabling data labels with a transparent background, disabling automatic shape resizing, manually setting label width and height in pixels, and customizing font color and size before saving the workbook.
    public class ResizeDataLabelShapesDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
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

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the series
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the first series
            Series series = chart.NSeries[0];

            // Enable data labels and show the values
            series.DataLabels.ShowValue = true;

            // Set a transparent background to test contrast
            series.DataLabels.BackgroundMode = BackgroundMode.Transparent;

            // Disable automatic resizing of the shape to fit the text
            series.DataLabels.IsResizeShapeToFitText = false;

            // Manually set the size of the data label shape (in pixels)
            series.DataLabels.WidthPixel = 80;   // custom width
            series.DataLabels.HeightPixel = 30;  // custom height

            // Optionally change the font color to see contrast against the transparent background
            series.DataLabels.Font.Color = Color.Black;
            series.DataLabels.Font.Size = 12;

            // Save the workbook
            workbook.Save("ResizeDataLabelShapesDemo.xlsx");
        }
    }
}
