// Title: Automatic Scaling of the Secondary Axis in an Aspose.Cells Column Chart (C#)
// Description: Demonstrates how to create a workbook with primary and secondary data series, add a column chart, plot the secondary series on a secondary value axis, and enable automatic calculation of the axis minimum and maximum values using Aspose.Cells for .NET.
// Keywords: Aspose.Cells secondary axis auto scaling | C# Aspose.Cells chart secondary axis | automatic min max secondary axis Aspose.Cells | dual axis chart Aspose.Cells .NET | column chart secondary value axis Aspose.Cells | Aspose.Cells chart scaling | Excel secondary axis auto range C#
// Common Searches: Aspose.Cells set secondary axis automatic min max C# | enable auto scaling for secondary axis in Aspose.Cells chart | Aspose.Cells dual axis chart example C# | automatic secondary value axis range Aspose.Cells .NET | C# Aspose.Cells column chart with secondary axis auto scaling
// Developer Intent: Automatically determine the minimum and maximum limits for a chart's secondary value axis without manually specifying them.
// Use Cases: Create a dual‑axis column chart where the secondary series has a vastly different range and let the secondary axis adjust automatically. | Generate Excel reports that include charts with dynamic data ranges, ensuring the secondary axis scales correctly for each export. | Save workbooks with auto‑scaled secondary axes for downstream analysis, presentations, or further programmatic manipulation.
// AI Prompts: Show how to switch the chart to a line type while keeping automatic secondary axis scaling in Aspose.Cells C#. | Provide code to customize the secondary axis title font, color, and alignment after enabling auto scaling. | Explain how to disable automatic scaling and set explicit minimum and maximum values for the secondary axis in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook with primary and secondary data series, add a column chart, plot the secondary series on a secondary value axis, and enable automatic calculation of the axis minimum and maximum values using Aspose.Cells for .NET.
    public class SecondaryAxisAutoScalingDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");

                // Primary series values (smaller range)
                worksheet.Cells["B1"].PutValue("Primary");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);

                // Secondary series values (larger range)
                worksheet.Cells["C1"].PutValue("Secondary");
                worksheet.Cells["C2"].PutValue(1000);
                worksheet.Cells["C3"].PutValue(2000);
                worksheet.Cells["C4"].PutValue(3000);

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = worksheet.Charts[chartIndex];

                // Add the two series to the chart
                chart.NSeries.Add("B2:B4", true); // primary series
                chart.NSeries.Add("C2:C4", true); // secondary series
                chart.NSeries.CategoryData = "A2:A4";

                // Plot the second series on the secondary value axis
                chart.NSeries[1].PlotOnSecondAxis = true;

                // Access the secondary value axis
                Axis secondaryAxis = chart.SecondValueAxis;

                // Enable automatic calculation of min and max values
                secondaryAxis.IsAutomaticMinValue = true;
                secondaryAxis.IsAutomaticMaxValue = true;

                // Optional title for the secondary axis
                secondaryAxis.Title.Text = "Secondary Axis";

                // Determine output file path
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "SecondaryAxisAutoScalingDemo.xlsx");

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
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
