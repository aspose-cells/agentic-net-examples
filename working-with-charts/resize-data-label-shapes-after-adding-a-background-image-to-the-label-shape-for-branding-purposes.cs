// Title: Aspose.Cells for .NET – Resize Chart Data Label Shapes by Disabling Auto‑Size
// Description: Demonstrates how to create a workbook, add a column chart, enable data labels, turn off automatic shape resizing (IsResizeShapeToFitText = false), and assign a fixed width and height to each label shape—ideal for adding branded background images.
// Keywords: Aspose.Cells resize data label | chart data label custom size .NET | disable auto‑size data label Aspose | set label width height Aspose.Cells | branding chart data labels C# | fixed size data label shape
// Common Searches: how to set fixed size for chart data labels Aspose.Cells | prevent data label shape from auto‑resizing in .NET | resize data label shapes after adding background image | custom width and height for Excel chart labels C#
// Developer Intent: Apply a consistent, fixed dimension to every chart data label shape so branding graphics align correctly.
// Use Cases: Generate a column chart where each data label has an 80 × 30 pixel shape for uniform logo placement. | Iterate through all points in a series, disable auto‑size, and set custom dimensions before saving the workbook. | Prepare Excel reports with branded data labels that require a predetermined shape size to match background images.
// AI Prompts: Show C# code to disable IsResizeShapeToFitText and set Width/Height for chart data labels using Aspose.Cells. | Explain how to resize data label shapes after adding a background image in an Excel chart with Aspose.Cells for .NET. | Provide a step‑by‑step guide to apply a fixed size to all data labels in a column chart without affecting other chart elements.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsBrandingExample
{
    // Demonstrates how to create a workbook, add a column chart, enable data labels, turn off automatic shape resizing (IsResizeShapeToFitText = false), and assign a fixed width and height to each label shape—ideal for adding branded background images.
    public class ResizeDataLabelShapes
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
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
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIdx];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the first series
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;
                series.DataLabels.Position = LabelPositionType.Center;

                // Iterate through each point and customize its data label
                foreach (ChartPoint point in series.Points)
                {
                    // Access the data label (ChartTextFrame)
                    DataLabels label = point.DataLabels;

                    // Prevent the shape from auto‑resizing to fit the text
                    label.IsResizeShapeToFitText = false;

                    // Define custom size for the label shape (in pixels)
                    // Width and Height are obsolete but still functional; warnings are acceptable
                    label.Width = 80;   // adjust as needed
                    label.Height = 30;  // adjust as needed
                }

                // Save the workbook
                string outputPath = "ResizedDataLabels.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ResizeDataLabelShapes.Run();
        }
    }
}
