// Title: Aspose.Cells C# – Auto‑Scale Secondary Axis in a Column Chart
// Description: Demonstrates how to create a workbook, add primary and secondary series, plot a column chart, and enable automatic minimum and maximum calculation for the secondary value axis using IsAutomaticMinValue and IsAutomaticMaxValue.
// Keywords: Aspose.Cells secondary axis auto scaling | C# chart automatic min max | IsAutomaticMinValue Aspose.Cells | IsAutomaticMaxValue Aspose.Cells | secondary value axis C# | Aspose.Cells column chart example | auto scale secondary axis .NET
// Common Searches: Aspose.Cells enable auto scaling for secondary axis C# | set secondary value axis automatic min max Aspose.Cells | C# chart secondary axis auto range Aspose.Cells | how to auto scale secondary axis in Aspose.Cells chart | Aspose.Cells secondary axis IsAutomaticMinValue example
// Developer Intent: Configure a chart’s secondary value axis to calculate its minimum and maximum limits automatically.
// Use Cases: Display sales volume and profit margin together, letting the margin axis auto‑scale to keep both series readable. | Generate financial dashboards where percentages and absolute values share a chart without manual axis adjustments. | Create reports that combine large‑scale data (e.g., units sold) with small‑scale metrics (e.g., growth rate) using automatic secondary axis scaling.
// AI Prompts: Write C# code with Aspose.Cells that adds a column chart and enables automatic scaling for the secondary axis. | Explain the effect of IsAutomaticMinValue and IsAutomaticMaxValue on a secondary axis and show how to set a custom axis title. | Provide a step‑by‑step tutorial for building a dual‑axis chart in Aspose.Cells .NET where the secondary axis auto‑determines its range.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add primary and secondary series, plot a column chart, and enable automatic minimum and maximum calculation for the secondary value axis using IsAutomaticMinValue and IsAutomaticMaxValue.
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
                // Primary series values
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["B1"].PutValue("Series 1");
                worksheet.Cells["B2"].PutValue(100);
                worksheet.Cells["B3"].PutValue(200);
                worksheet.Cells["B4"].PutValue(300);

                // Secondary series values (different magnitude)
                worksheet.Cells["C1"].PutValue("Series 2");
                worksheet.Cells["C2"].PutValue(5000);
                worksheet.Cells["C3"].PutValue(3000);
                worksheet.Cells["C4"].PutValue(1000);

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
                Chart chart = worksheet.Charts[chartIndex];

                // Add the two series
                chart.NSeries.Add("B2:B4", true); // primary series
                chart.NSeries.Add("C2:C4", true); // secondary series
                chart.NSeries.CategoryData = "A2:A4";

                // Plot the second series on the secondary value axis
                chart.NSeries[1].PlotOnSecondAxis = true;

                // Access the secondary value axis
                Axis secondaryAxis = chart.SecondValueAxis;

                // Enable automatic calculation of minimum and maximum values
                secondaryAxis.IsAutomaticMinValue = true;
                secondaryAxis.IsAutomaticMaxValue = true;

                // (Optional) Set a title to identify the secondary axis
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
