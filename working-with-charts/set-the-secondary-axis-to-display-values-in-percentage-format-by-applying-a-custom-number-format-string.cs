// Title: Aspose.Cells C# Example – Format Secondary Axis Tick Labels as Percentages (0.00%)
// Description: Demonstrates how to create a workbook, add primary and secondary data series, plot the secondary series on a secondary value axis, and apply a custom number format ("0.00%") to the secondary axis tick labels in an Aspose.Cells column chart using C#.
// Keywords: Aspose.Cells secondary axis | C# chart percentage format | custom number format Aspose.Cells | secondary value axis tick labels | column chart Aspose.Cells example | percentage axis formatting | Aspose.Cells chart tutorial | C# Excel chart formatting
// Common Searches: Aspose.Cells set secondary axis to percentage | C# chart secondary axis custom number format | how to format secondary axis tick labels Aspose.Cells | apply 0.00% format to secondary axis in Excel using Aspose | Aspose.Cells column chart percentage axis example
// Developer Intent: The developer needs to display secondary axis values as percentages by applying a custom number format string to the axis tick labels in an Aspose.Cells chart.
// Use Cases: Financial dashboards that show absolute sales on the primary axis and growth rates as percentages on the secondary axis. | Performance reports where raw counts are plotted alongside conversion rates formatted as percentages. | Business intelligence visualizations that compare unit volumes with percentage change trends in a single chart.
// AI Prompts: Show C# code to set a secondary axis tick label format to "0.00%" in an Aspose.Cells column chart. | Explain how to plot a series on the secondary axis and apply a percentage number format using Aspose.Cells. | Provide a step‑by‑step Aspose.Cells example that formats the secondary value axis as percentages and adds axis titles.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add primary and secondary data series, plot the secondary series on a secondary value axis, and apply a custom number format ("0.00%") to the secondary axis tick labels in an Aspose.Cells column chart using C#.
    public class SecondaryAxisPercentageFormat
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

                // Primary series values
                worksheet.Cells["B1"].PutValue("Primary");
                worksheet.Cells["B2"].PutValue(100);
                worksheet.Cells["B3"].PutValue(200);
                worksheet.Cells["B4"].PutValue(300);

                // Secondary series values (as fractions to be shown as percentages)
                worksheet.Cells["C1"].PutValue("Secondary");
                worksheet.Cells["C2"].PutValue(0.25);
                worksheet.Cells["C3"].PutValue(0.5);
                worksheet.Cells["C4"].PutValue(0.75);

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
                Chart chart = worksheet.Charts[chartIndex];

                // Add the two data series
                chart.NSeries.Add("B2:B4", true); // primary series
                chart.NSeries.Add("C2:C4", true); // secondary series
                chart.NSeries.CategoryData = "A2:A4";

                // Plot the second series on the secondary value axis
                chart.NSeries[1].PlotOnSecondAxis = true;

                // Apply a custom number format to the secondary value axis tick labels (percentage)
                chart.SecondValueAxis.TickLabels.NumberFormat = "0.00%";

                // Optional: give titles to axes for clarity
                chart.ValueAxis.Title.Text = "Primary Axis";
                chart.SecondValueAxis.Title.Text = "Secondary Axis (Percentage)";

                // Save the workbook
                string outputPath = "SecondaryAxisPercentageFormat.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main()
        {
            Run();
        }
    }
}
