// Title: C# – Resize Bubble Chart Data Labels After Increasing BubbleScale with Aspose.Cells
// Description: This example creates a workbook, adds a bubble chart, sets BubbleScale to 200%, enables data labels, calculates the chart to obtain each point's pixel radius, disables automatic label resizing, and assigns custom Width and Height based on the bubble's diameter (converted from pixels to points). The workbook is then saved as an Excel file.
// Keywords: Aspose.Cells C# bubble chart | resize data labels Aspose.Cells | BubbleScale .NET | ChartPoint RadiusPx conversion | custom label size Excel chart | disable automatic label resizing | pixel to point conversion chart
// Common Searches: how to change bubble chart data label size Aspose.Cells | increase bubble scale and adjust labels C# | set custom width height for chart point labels | Aspose.Cells resize data labels after scaling bubbles | convert pixel radius to points for Excel chart labels
// Developer Intent: Programmatically set each bubble's data label dimensions to match the scaled bubble size.
// Use Cases: Maintain proportional data labels when bubble sizes are enlarged for better visual balance. | Generate reports with bubble charts where labels need precise sizing to avoid overlap. | Dynamically compute label dimensions from a point's RadiusPx for responsive chart rendering.
// AI Prompts: Write C# code using Aspose.Cells that increases BubbleScale to 200% and resizes each bubble's data label based on ChartPoint.RadiusPx. | Show how to convert a bubble's pixel radius to points and apply the values to DataLabels.Width and DataLabels.Height in a bubble chart. | Provide an Aspose.Cells example that disables automatic label resizing and sets custom label dimensions for every point after scaling bubbles.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This example creates a workbook, adds a bubble chart, sets BubbleScale to 200%, enables data labels, calculates the chart to obtain each point's pixel radius, disables automatic label resizing, and assigns custom Width and Height based on the bubble's diameter (converted from pixels to points). The workbook is then saved as an Excel file.
    public class ResizeDataLabelShapesForBubbleChart
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

            // Populate sample data for X, Y and bubble size
            worksheet.Cells["A1"].PutValue("X");
            worksheet.Cells["B1"].PutValue("Y");
            worksheet.Cells["C1"].PutValue("Size");

            worksheet.Cells["A2"].PutValue(1);
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["C2"].PutValue(5);

            worksheet.Cells["A3"].PutValue(2);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["C3"].PutValue(10);

            worksheet.Cells["A4"].PutValue(3);
            worksheet.Cells["B4"].PutValue(30);
            worksheet.Cells["C4"].PutValue(15);

            // Add a bubble chart
            int chartIndex = worksheet.Charts.Add(ChartType.Bubble, 5, 0, 20, 10);
            Chart chart = worksheet.Charts[chartIndex];

            // Add series and bind data
            int seriesIndex = chart.NSeries.Add("B2:B4", true); // Y values
            chart.NSeries[seriesIndex].XValues = "A2:A4";      // X values
            chart.NSeries[seriesIndex].BubbleSizes = "C2:C4"; // Bubble sizes

            // Increase bubble size scaling to test label resizing
            chart.NSeries[seriesIndex].BubbleScale = 200; // 200% of default size

            // Enable data labels to be visible
            Series series = chart.NSeries[seriesIndex];
            series.DataLabels.ShowValue = true;
            series.DataLabels.ShowBubbleSize = true;

            // Calculate the chart so that RadiusPx values are populated
            chart.Calculate();

            // Iterate over each point and resize its data label shape
            foreach (ChartPoint point in series.Points)
            {
                // Disable automatic resizing of the shape to fit text
                point.DataLabels.IsResizeShapeToFitText = false;

                // Use the bubble radius (in pixels) to set a custom width/height for the label
                int diameter = point.RadiusPx * 2;

                // Convert pixels to points (1 point = 1/72 inch, assuming 96 DPI)
                double points = diameter * 72.0 / 96.0;

                point.DataLabels.Width = (int)points;
                point.DataLabels.Height = (int)points;
            }

            // Save the workbook
            string outputPath = "ResizeDataLabelShapesBubbleChart.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
