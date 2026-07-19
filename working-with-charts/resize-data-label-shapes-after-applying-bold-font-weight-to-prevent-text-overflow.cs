// Title: Resize Chart Data Label Shapes After Applying Bold Font with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a column chart in Aspose.Cells, enable data labels, apply bold formatting, and then resize each label shape by setting WidthPixel and HeightPixel to prevent text overflow before saving the workbook.
// Keywords: Aspose.Cells resize data label | chart data label bold font .NET | prevent label overflow Excel | ChartPoint DataLabels WidthPixel | IsResizeShapeToFitText Aspose | C# Aspose.Cells chart label sizing | Excel chart label custom dimensions
// Common Searches: how to increase chart data label size after bold in Aspose.Cells | Aspose.Cells .NET prevent data label clipping | set custom width and height for Excel chart labels using C# | disable automatic resizing of chart data labels Aspose
// Developer Intent: Adjust the dimensions of chart data label shapes after applying bold styling to avoid text clipping.
// Use Cases: Generate a column chart where bold data labels remain fully visible by disabling auto‑resize and defining fixed pixel dimensions. | Export Excel workbooks with consistently sized data labels across multiple series after changing font weight. | Create reports that require bold emphasis on values without risking label overflow or layout distortion.
// AI Prompts: Write C# code with Aspose.Cells that makes chart data labels bold and then sets a fixed WidthPixel and HeightPixel for each label to stop overflow. | Show an Aspose.Cells .NET example that disables IsResizeShapeToFitText for ChartPoint data labels and assigns specific pixel sizes after changing the font style. | Explain how to iterate over ChartPoint objects in a chart to adjust DataLabels properties such as IsResizeShapeToFitText, WidthPixel, and HeightPixel.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a column chart in Aspose.Cells, enable data labels, apply bold formatting, and then resize each label shape by setting WidthPixel and HeightPixel to prevent text overflow before saving the workbook.
    public class ResizeDataLabelShapesAfterBoldDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the first series
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;

                // Apply bold font to all data labels
                series.DataLabels.Font.IsBold = true;
                // Ensure the font change is applied to each individual label
                series.DataLabels.ApplyFont();

                // Resize each data label shape to avoid overflow caused by bold text
                foreach (ChartPoint point in series.Points)
                {
                    // Disable automatic resizing so we can set a custom size
                    point.DataLabels.IsResizeShapeToFitText = false;

                    // Increase the width (in pixels) to give more room for the bold text
                    point.DataLabels.WidthPixel = 80;

                    // Optionally increase height as well
                    point.DataLabels.HeightPixel = 30;
                }

                // Save the workbook
                workbook.Save("ResizeDataLabelShapesAfterBoldDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ResizeDataLabelShapesAfterBoldDemo.Run();
        }
    }
}
