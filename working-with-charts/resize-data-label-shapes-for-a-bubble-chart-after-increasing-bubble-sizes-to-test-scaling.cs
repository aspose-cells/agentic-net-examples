// Title: Resize Bubble Chart Data Labels After Scaling Bubbles with Aspose.Cells for .NET
// Description: Demonstrates how to create a bubble chart, increase its BubbleScale to 200%, enable value and size data labels, calculate the chart, and then set a fixed width and height for each data label shape using Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# bubble chart | BubbleScale | resize data label shape | fixed label width height | disable auto‑fit data label | chart.Calculate() | .NET Excel chart | global | USA | UK | India
// Common Searches: how to change bubble chart data label size Aspose.Cells | set fixed width and height for chart data labels .NET | increase bubble scale without stretching labels | disable auto fit for Excel chart data labels using Aspose | resize data label shapes after chart.Calculate
// Developer Intent: Set a custom size for bubble chart data label shapes after adjusting the bubble scaling.
// Use Cases: Generate an Excel workbook with a bubble chart where bubbles are enlarged (200% scale) but labels keep a consistent 60 × 30‑point size for printing. | Create dashboards that show both the data value and bubble size while preventing label auto‑fit from altering layout. | Perform post‑calculation layout adjustments on chart points to align labels with other UI elements.
// AI Prompts: Provide C# code that sets BubbleScale on a bubble chart series and then fixes each data label shape to a specific width and height with Aspose.Cells. | Explain the steps to disable auto‑fit for bubble chart data labels and apply uniform dimensions after calling chart.Calculate(). | Show how to iterate over ChartPoint objects to modify DataLabels properties such as IsResizeShapeToFitText, Width, and Height.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a bubble chart, increase its BubbleScale to 200%, enable value and size data labels, calculate the chart, and then set a fixed width and height for each data label shape using Aspose.Cells for C#.
    public class ResizeBubbleChartDataLabels
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
                for (int i = 2; i <= 5; i++)
                {
                    sheet.Cells[$"A{i}"].PutValue(i);          // X values
                    sheet.Cells[$"B{i}"].PutValue(i * 10);     // Y values
                    sheet.Cells[$"C{i}"].PutValue(i * 2);      // Bubble size values
                }

                // Add a bubble chart
                int chartIndex = sheet.Charts.Add(ChartType.Bubble, 5, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Set series data
                int seriesIndex = chart.NSeries.Add("B2:B5", true); // Y values
                chart.NSeries[seriesIndex].XValues = "A2:A5";       // X values
                chart.NSeries[seriesIndex].BubbleSizes = "C2:C5";   // Size values

                // Increase bubble sizes to test scaling (e.g., 200% of default)
                chart.NSeries[seriesIndex].BubbleScale = 200;

                // Show data labels (value and bubble size) for each point
                Series series = chart.NSeries[seriesIndex];
                series.DataLabels.ShowValue = true;
                series.DataLabels.ShowBubbleSize = true;

                // Calculate the chart so that runtime properties (like RadiusPx) are populated
                chart.Calculate();

                // Resize each data label shape: disable auto‑fit and set a fixed width/height
                foreach (ChartPoint point in series.Points)
                {
                    point.DataLabels.IsResizeShapeToFitText = false;
                    point.DataLabels.Width = 60;   // width in points
                    point.DataLabels.Height = 30;  // height in points
                }

                // Save the workbook with the modified chart
                workbook.Save("ResizedBubbleChartDataLabels.xlsx");
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
            ResizeBubbleChartDataLabels.Run();
        }
    }
}
