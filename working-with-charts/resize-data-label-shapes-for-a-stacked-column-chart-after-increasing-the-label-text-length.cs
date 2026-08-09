// Title: Resize data label shapes in a stacked column chart using Aspose.Cells for .NET
// Description: Demonstrates how to create a stacked column chart, extend each data label's text, turn off automatic resizing, and assign a fixed width (120 px) and height (30 px) to every label before recalculating and saving the workbook.
// Keywords: Aspose.Cells | C# chart data labels | stacked column chart | custom label size | disable auto resize | IsResizeShapeToFitText | .NET Excel chart | ChartPoint label text | LabelPositionType.Center | set label width height
// Common Searches: Aspose.Cells set fixed size for chart data labels | how to prevent data label auto‑resize in .NET | increase data label text and adjust shape Aspose.Cells | customize stacked column chart labels C# | resize data label shape for each point Aspose.Cells
// Developer Intent: Apply a consistent, manually defined size to data labels after lengthening their text in a stacked column chart.
// Use Cases: Ensuring readability when data labels contain additional descriptive text. | Maintaining uniform label dimensions across all points in a series. | Creating printable Excel reports where label size must not change automatically.
// AI Prompts: Write C# code with Aspose.Cells that disables automatic label resizing and sets a fixed width and height for each data label in a stacked column chart. | Show how to append extra information to chart point labels and then resize the label shape to fit the new text. | Provide an example of customizing data label position and size for a specific series in an Aspose.Cells-generated Excel chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsResizeDataLabelDemo
{
    // Demonstrates how to create a stacked column chart, extend each data label's text, turn off automatic resizing, and assign a fixed width (120 px) and height (30 px) to every label before recalculating and saving the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a stacked column chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");

                // First series (Product A)
                sheet.Cells["B1"].PutValue("Product A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Second series (Product B)
                sheet.Cells["C1"].PutValue("Product B");
                sheet.Cells["C2"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["C4"].PutValue(35);

                // Add a stacked column chart (ColumnStacked is the correct enum value)
                int chartIdx = sheet.Charts.Add(ChartType.ColumnStacked, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIdx];

                // Set data range for both series
                chart.NSeries.Add("B2:C4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the first series
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;
                series.DataLabels.Position = LabelPositionType.Center;

                // Increase label text length for each point and resize the label shape
                foreach (ChartPoint point in series.Points)
                {
                    // Append extra text to make the label longer
                    point.DataLabels.Text = $"Value: {point.YValue} (extended info)";

                    // Disable automatic resizing so we can set custom dimensions
                    point.DataLabels.IsResizeShapeToFitText = false;

                    // Set custom size (in pixels) sufficient for the longer text
                    point.DataLabels.Width = 120;
                    point.DataLabels.Height = 30;
                }

                // Recalculate the chart to apply changes
                chart.Calculate();

                // Save the workbook
                string outputPath = "StackedColumn_ResizedDataLabels.xlsx";
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
