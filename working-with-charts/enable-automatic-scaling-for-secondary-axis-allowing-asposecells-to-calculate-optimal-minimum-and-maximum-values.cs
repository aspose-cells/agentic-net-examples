// Title: Aspose.Cells for .NET – Automatic Scaling of a Chart’s Secondary Axis
// Description: Demonstrates how to create a column chart with two series, plot the second series on the secondary value axis, and let Aspose.Cells automatically calculate the optimal minimum and maximum values by enabling IsAutomaticMinValue and IsAutomaticMaxValue.
// Keywords: Aspose.Cells secondary axis auto scaling | C# chart automatic min max | Aspose.Cells secondary value axis | column chart secondary axis | Aspose.Cells axis configuration | auto range secondary axis | chart scaling Aspose.Cells
// Common Searches: Aspose.Cells enable automatic scaling secondary axis C# | set secondary axis auto min max Aspose.Cells | C# chart secondary value axis auto range example | Aspose.Cells column chart secondary axis scaling | how to auto scale secondary axis in Aspose.Cells
// Developer Intent: Configure a chart so that the secondary value axis automatically determines its minimum and maximum limits based on the plotted data.
// Use Cases: Financial reports that show revenue on the primary axis and profit on a secondary axis without manual range settings. | Dashboards combining metrics with vastly different scales, letting Aspose.Cells compute appropriate bounds for the secondary axis. | Scientific charts where a secondary measurement unit requires dynamic scaling alongside the primary data series.
// AI Prompts: Show me C# code to enable automatic scaling for a chart’s secondary axis using Aspose.Cells. | How do I set IsAutomaticMinValue and IsAutomaticMaxValue for the secondary value axis in Aspose.Cells? | Explain the algorithm Aspose.Cells uses to calculate optimal axis limits when auto‑scaling is turned on.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a column chart with two series, plot the second series on the secondary value axis, and let Aspose.Cells automatically calculate the optimal minimum and maximum values by enabling IsAutomaticMinValue and IsAutomaticMaxValue.
    public class SecondaryAxisAutoScalingDemo
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

                // First series (plotted on primary axis)
                sheet.Cells["B1"].PutValue("Series 1");
                sheet.Cells["B2"].PutValue(100);
                sheet.Cells["B3"].PutValue(200);
                sheet.Cells["B4"].PutValue(300);

                // Second series (will be plotted on secondary axis)
                sheet.Cells["C1"].PutValue("Series 2");
                sheet.Cells["C2"].PutValue(5000);
                sheet.Cells["C3"].PutValue(3000);
                sheet.Cells["C4"].PutValue(1000);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Add the two series to the chart
                chart.NSeries.Add("B2:B4", true);   // Series 1
                chart.NSeries.Add("C2:C4", true);   // Series 2
                chart.NSeries.CategoryData = "A2:A4";

                // Plot the second series on the secondary value axis
                chart.NSeries[1].PlotOnSecondAxis = true;

                // Enable automatic scaling for the secondary value axis
                Axis secondaryAxis = chart.SecondValueAxis;
                secondaryAxis.IsAutomaticMinValue = true; // Let Aspose.Cells calculate optimal minimum
                secondaryAxis.IsAutomaticMaxValue = true; // Let Aspose.Cells calculate optimal maximum

                // (Optional) Give the secondary axis a title for clarity
                secondaryAxis.Title.Text = "Secondary Axis";

                // Save the workbook
                string outputPath = "SecondaryAxisAutoScalingDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
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
            SecondaryAxisAutoScalingDemo.Run();
        }
    }
}
