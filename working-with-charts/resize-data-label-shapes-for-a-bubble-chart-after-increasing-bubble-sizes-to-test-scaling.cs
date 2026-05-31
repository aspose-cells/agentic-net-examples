using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class ResizeDataLabelShapesForBubbleChart
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for X, Y and bubble size
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
                int chartIndex = sheet.Charts.Add(ChartType.Bubble, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Add series and bind data
                int seriesIndex = chart.NSeries.Add("B2:B4", true); // Y values
                chart.NSeries[seriesIndex].XValues = "A2:A4";      // X values
                chart.NSeries[seriesIndex].BubbleSizes = "C2:C4"; // Bubble sizes

                // Increase bubble sizes using BubbleScale (e.g., 200% of default)
                chart.NSeries[seriesIndex].BubbleScale = 200;

                // Enable data labels so we can resize their shapes
                chart.NSeries[seriesIndex].DataLabels.ShowValue = true;

                // Iterate through each point and adjust its data label shape
                foreach (ChartPoint point in chart.NSeries[seriesIndex].Points)
                {
                    // Disable auto‑fit of the shape to the text
                    point.DataLabels.IsResizeShapeToFitText = false;

                    // Set a custom width and height (units are points)
                    point.DataLabels.Width = 60;   // narrower than default
                    point.DataLabels.Height = 30; // shorter than default
                }

                // Recalculate the chart to apply changes
                chart.Calculate();

                // Save the workbook
                string outputPath = "ResizeDataLabelShapesBubbleChart.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ResizeDataLabelShapesForBubbleChart.Run();
        }
    }
}