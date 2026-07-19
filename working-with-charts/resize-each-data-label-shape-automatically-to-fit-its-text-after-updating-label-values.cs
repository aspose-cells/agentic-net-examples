// Title: Auto‑Resize Chart Data Label Shapes to Fit Text with Aspose.Cells for .NET
// Description: Shows how to build a column chart, enable data labels, prepend custom text to each label, and activate the IsResizeShapeToFitText flag so every label shape automatically expands to contain the new text before saving the workbook.
// Keywords: Aspose.Cells | C# | chart data labels | auto resize | IsResizeShapeToFitText | column chart | Excel automation | dynamic label text | shape fit text | Aspose.Cells for .NET
// Common Searches: Aspose.Cells resize data label shape | C# auto fit chart label text | IsResizeShapeToFitText example | how to adjust chart data label size Aspose.Cells | auto resize Excel chart labels C#
// Developer Intent: Resize each chart data label shape so it automatically fits the updated label text.
// Use Cases: Add a prefix or suffix to data labels and ensure the label box grows to show the full string without clipping. | Create Excel reports with column charts where label lengths vary, maintaining a clean layout through automatic shape resizing. | Generate dynamic dashboards where label content changes at runtime and the chart adapts without manual size adjustments.
// AI Prompts: Write C# code using Aspose.Cells to update chart data label text and enable automatic shape resizing for every point. | Explain the behavior and limitations of the IsResizeShapeToFitText property when resizing data label shapes in Aspose.Cells. | Show how to apply auto‑resize to data labels across multiple series in a single chart using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Shows how to build a column chart, enable data labels, prepend custom text to each label, and activate the IsResizeShapeToFitText flag so every label shape automatically expands to contain the new text before saving the workbook.
    public class ResizeDataLabelShapesDemo
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
                sheet.Cells["A2"].PutValue("Alpha");
                sheet.Cells["A3"].PutValue("Beta");
                sheet.Cells["A4"].PutValue("Gamma");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(123);
                sheet.Cells["B3"].PutValue(4567);
                sheet.Cells["B4"].PutValue(89);

                // Add a column chart
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 15);
                Chart chart = sheet.Charts[chartIdx];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the first series
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;
                series.DataLabels.Position = LabelPositionType.Center;

                // Update each data label's text and enable auto‑resize to fit the text
                foreach (ChartPoint point in series.Points)
                {
                    // Prepend a custom prefix to the existing value
                    point.DataLabels.Text = $"Val: {point.YValue}";

                    // Ensure the shape automatically resizes to contain the new text
                    point.DataLabels.IsResizeShapeToFitText = true;
                }

                // Save the workbook
                workbook.Save("ResizeDataLabelShapesDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ResizeDataLabelShapesDemo.Run();
        }
    }
}
