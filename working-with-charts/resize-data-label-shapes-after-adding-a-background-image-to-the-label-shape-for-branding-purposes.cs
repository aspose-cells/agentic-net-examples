// Title: Resize Chart Data Label Shapes for Branding with Aspose.Cells for .NET
// Description: Demonstrates how to create a column chart, enable data labels, disable automatic shape resizing, and set a fixed width (60 px) and height (30 px) for each label so a background image can be applied without distortion.
// Keywords: Aspose.Cells chart data label size | C# resize data label shape | disable IsResizeShapeToFitText | fixed width height chart label | add background image to data label | .NET Excel chart customization
// Common Searches: set fixed size for Excel chart data labels Aspose.Cells | prevent data label auto‑resize after adding image | customize chart label dimensions C# | Aspose.Cells label background image sizing | how to control data label shape size in .NET
// Developer Intent: Control the dimensions of chart data label shapes after applying a branding background image, ensuring consistent appearance across all labels.
// Use Cases: Create a column chart and assign each data label a 60 × 30 pixel rectangle for logo placement. | Turn off IsResizeShapeToFitText to keep label shapes from expanding with longer text. | Maintain uniform label size when exporting Excel files to PDF or image formats.
// AI Prompts: Write C# code using Aspose.Cells that adds a background image to each chart data label and then fixes the label Width to 60 and Height to 30 pixels. | Provide an example that disables IsResizeShapeToFitText for all points in a series, sets custom font color, and ensures the label shape size remains constant.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsLabelResizeDemo
{
    // Demonstrates how to create a column chart, enable data labels, disable automatic shape resizing, and set a fixed width (60 px) and height (30 px) for each label so a background image can be applied without distortion.
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
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIdx];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the first series
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;
                series.DataLabels.Position = LabelPositionType.Center;

                // Customize each data label
                foreach (ChartPoint point in series.Points)
                {
                    // Disable automatic resizing so we can control the shape size
                    point.DataLabels.IsResizeShapeToFitText = false;

                    // Set a custom size (smaller than the default needed for the text)
                    point.DataLabels.Width = 60;   // pixels
                    point.DataLabels.Height = 30; // pixels

                    // Optional: set font color or other properties if needed
                    point.DataLabels.Font.Color = Color.Black;
                }

                // Save the workbook
                workbook.Save("ChartDataLabelsResized.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
