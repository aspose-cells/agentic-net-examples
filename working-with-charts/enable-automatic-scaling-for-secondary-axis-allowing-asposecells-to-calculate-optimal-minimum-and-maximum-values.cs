// Title: Enable automatic min/max scaling for the secondary value axis in an Aspose.Cells column chart using C#
// AI Prompts: Generate C# code that creates a column chart with primary and secondary series, plots the secondary series on the secondary axis, and enables the axis to compute its own minimum and maximum values using Aspose.Cells. | Show how to activate the IsAutomaticMinValue and IsAutomaticMaxValue flags for a chart’s secondary value axis in a .NET workbook.
// Common Searches: Aspose.Cells C# set secondary axis auto range for column chart | How to turn on auto range for secondary value axis in Aspose.Cells | C# chart secondary axis min max auto calculation Aspose.Cells example | Enable auto min and max values on secondary axis using Aspose.Cells .NET | Aspose.Cells column chart with secondary axis auto range tutorial
// Tags: secondary value axis auto scaling Aspose.Cells | C# column chart secondary axis Aspose.Cells | IsAutomaticMinValue property Aspose.Cells | IsAutomaticMaxValue property Aspose.Cells | Aspose.Cells chart axis auto range

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a workbook, adds sample data, builds a column chart with primary and secondary series, assigns the second series to the secondary value axis, turns on automatic minimum and maximum calculation for that axis, optionally sets a title, and saves the workbook as an XLSX file.
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
                worksheet.Cells["B1"].PutValue("Series 1");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);

                // Secondary series values (larger range)
                worksheet.Cells["C1"].PutValue("Series 2");
                worksheet.Cells["C2"].PutValue(5000);
                worksheet.Cells["C3"].PutValue(3000);
                worksheet.Cells["C4"].PutValue(1000);

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = worksheet.Charts[chartIndex];

                // Add the two series to the chart
                chart.NSeries.Add("B2:B4", true); // Series 1
                chart.NSeries.Add("C2:C4", true); // Series 2
                chart.NSeries.CategoryData = "A2:A4";

                // Plot the second series on the secondary value axis
                chart.NSeries[1].PlotOnSecondAxis = true;

                // Access the secondary value axis
                Axis secondaryAxis = chart.SecondValueAxis;

                // Enable automatic calculation of min and max values
                secondaryAxis.IsAutomaticMinValue = true;
                secondaryAxis.IsAutomaticMaxValue = true;

                // (Optional) Set a title for clarity
                secondaryAxis.Title.Text = "Secondary Axis (Auto-Scaled)";

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

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
