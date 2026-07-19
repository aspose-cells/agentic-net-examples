// Title: Resize Bubble Chart Data Label Shapes After Scaling Bubbles – Aspose.Cells for .NET
// Description: Shows how to increase a bubble chart's BubbleScale, turn off auto‑fit, and assign explicit Width and Height to each data label shape using Aspose.Cells in C#.
// Keywords: Aspose.Cells | C# | bubble chart | BubbleScale | data label size | custom label dimensions | disable auto‑fit | chart point | Excel automation | resize data labels
// Common Searches: Aspose.Cells set bubble scale | change data label size in bubble chart Aspose | disable data label auto resize Aspose.Cells | custom width height for chart data labels .NET | resize bubble chart labels after scaling
// Developer Intent: Set fixed width and height for bubble chart data label shapes after increasing bubble scaling.
// Use Cases: Maintain consistent label dimensions when bubble sizes are enlarged to avoid overlap. | Create Excel reports where data labels stay uniform regardless of bubble scale adjustments. | Fine‑tune chart layout for presentations that require precise label placement.
// AI Prompts: Provide C# code that disables DataLabels.IsResizeShapeToFitText and sets Width/Height for each ChartPoint in an Aspose.Cells bubble chart. | Show an example of increasing BubbleScale to 200% while keeping data label shapes at 80 × 30 points using Aspose.Cells for .NET. | Explain how to recalculate a chart after modifying data label shape sizes to ensure correct positioning in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Shows how to increase a bubble chart's BubbleScale, turn off auto‑fit, and assign explicit Width and Height to each data label shape using Aspose.Cells in C#.
    public class ResizeDataLabelShapesForBubbleChart
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a bubble chart (X, Y, Size)
                sheet.Cells["A1"].PutValue("X");
                sheet.Cells["B1"].PutValue("Y");
                sheet.Cells["C1"].PutValue("Size");

                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["C2"].PutValue(5);

                sheet.Cells["A3"].PutValue(2);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["C3"].PutValue(10);

                sheet.Cells["A4"].PutValue(3);
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["C4"].PutValue(15);

                // Add a bubble chart
                int chartIndex = sheet.Charts.Add(ChartType.Bubble, 5, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Add series and bind data
                int seriesIndex = chart.NSeries.Add("B2:B4", true); // Y values
                chart.NSeries[seriesIndex].XValues = "A2:A4";      // X values
                chart.NSeries[seriesIndex].BubbleSizes = "C2:C4"; // Bubble sizes

                // Increase bubble size scaling to test label scaling (e.g., 200%)
                chart.NSeries[seriesIndex].BubbleScale = 200;

                // Enable data labels and show both value and bubble size
                Series series = chart.NSeries[seriesIndex];
                series.DataLabels.ShowValue = true;
                series.DataLabels.ShowBubbleSize = true;

                // For each point, disable auto‑fit and set a custom width/height for the label shape
                foreach (ChartPoint point in series.Points)
                {
                    // Prevent the shape from automatically resizing to fit the text
                    point.DataLabels.IsResizeShapeToFitText = false;

                    // Set explicit dimensions (units are points)
                    point.DataLabels.Width = 80;   // narrower than default
                    point.DataLabels.Height = 30; // shorter than default
                }

                // Recalculate the chart so that all layout information (including label positions) is updated
                chart.Calculate();

                // Save the workbook
                workbook.Save("ResizeDataLabelShapesBubbleChart.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ResizeDataLabelShapesForBubbleChart.Run();
        }
    }
}
