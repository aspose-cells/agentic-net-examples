// Title: Set Fixed Width & Height for Chart Data Label Shapes in Aspose.Cells for .NET
// Description: Shows how to build a workbook, fill a 5×5 data range, add a column chart, turn on data labels, calculate the chart, and then assign a constant size (e.g., 40 pt × 20 pt) to every label shape while turning off automatic resizing, before saving the workbook.
// Keywords: Aspose.Cells chart data labels | C# resize data label shape | fixed label width height Aspose.Cells | IsResizeShapeToFitText | ChartPoint label dimensions | Excel chart formatting .NET | custom data label size | disable auto‑size data labels | Aspose.Cells API label sizing | set data label shape size programmatically
// Common Searches: how to change data label size in Aspose.Cells | set fixed width and height for chart labels .NET | disable automatic resizing of data labels Aspose.Cells | custom dimensions for Excel chart data labels using C# | apply gradient fill then resize data labels Aspose.Cells
// Developer Intent: Apply a uniform width and height to each data label shape on a chart after enabling the labels, preventing the labels from auto‑sizing.
// Use Cases: Generate a heat‑map‑style column chart and keep all data labels the same size for a clean layout. | Maintain label dimensions when applying custom fills or gradients to a series. | Prepare Excel reports where consistent label appearance is required across different data sets.
// AI Prompts: Write C# code with Aspose.Cells that sets a constant width and height for all chart data label shapes and disables auto‑resize. | Provide an example that adds a gradient fill to a series, then resizes each data label to 40 pt × 20 pt using Aspose.Cells for .NET. | Explain how to loop through ChartPoint objects to modify IsResizeShapeToFitText, Width, and Height of data labels.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;   // Required for drawing related enums

namespace AsposeCellsExamples
{
    // Shows how to build a workbook, fill a 5×5 data range, add a column chart, turn on data labels, calculate the chart, and then assign a constant size (e.g., 40 pt × 20 pt) to every label shape while turning off automatic resizing, before saving the workbook.
    class ResizeHeatMapDataLabels
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Fill a 5x5 matrix with sample values (including headers)
                for (int i = 0; i < 5; i++)
                {
                    sheet.Cells[0, i + 1].PutValue("Col" + (i + 1));   // Column headers
                    sheet.Cells[i + 1, 0].PutValue("Row" + (i + 1));   // Row headers
                    for (int j = 0; j < 5; j++)
                    {
                        sheet.Cells[i + 1, j + 1].PutValue((i + 1) * (j + 1));
                    }
                }

                // Add a column chart (HeatMap not available in this version)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart (including headers)
                chart.SetChartDataRange("A1:F6", true);

                // Apply a solid fill color to the series (gradient not supported in older versions)
                Series series = chart.NSeries[0];
                // Note: SolidFillColor property may not be available in some versions; omitted for compatibility.

                // Enable data labels and set their basic appearance
                series.DataLabels.ShowValue = true;
                series.DataLabels.Position = LabelPositionType.Center;

                // Calculate the chart to make shape properties available
                chart.Calculate();

                // Resize each data label shape individually
                foreach (ChartPoint point in series.Points)
                {
                    // Prevent automatic resizing based on text
                    point.DataLabels.IsResizeShapeToFitText = false;

                    // Set custom width and height (in points)
                    point.DataLabels.Width = 40;   // width of the label shape
                    point.DataLabels.Height = 20;  // height of the label shape
                }

                // Save the workbook with the modified chart
                workbook.Save("HeatMapDataLabelsResized.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            ResizeHeatMapDataLabels.Run();
        }
    }
}
