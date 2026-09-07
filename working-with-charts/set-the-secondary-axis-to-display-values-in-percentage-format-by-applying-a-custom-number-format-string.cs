// Title: Apply a percentage custom number format to the secondary value axis of a column chart using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a column chart, adds a secondary value axis, and applies the "0.00%" number format to its tick labels using Aspose.Cells. | Generate an Aspose.Cells example that plots a second data series on a secondary axis and formats that axis to show percentages. | Show how to set a custom percentage format for the secondary axis tick labels in an Aspose.Cells chart with C#.
// Common Searches: Aspose.Cells C# format secondary chart axis as percentage | how to apply custom number format to secondary axis tick labels in Aspose.Cells | dual axis column chart percentage labels using Aspose.Cells .NET | C# Aspose.Cells set 0.00% number format for secondary axis
// Tags: secondary chart axis custom number format Aspose.Cells | percentage tick label formatting C# | dual axis column chart Aspose.Cells | custom number format string Aspose.Cells | chart axis formatting .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a workbook, adds sample data, builds a column chart with primary and secondary series, plots the second series on a secondary axis, applies the custom number format "0.00%" to the secondary axis tick labels, sets axis titles, and saves the file as SecondaryAxisPercentageFormat.xlsx.
    public class SecondaryAxisPercentageFormat
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully: SecondaryAxisPercentageFormat.xlsx");
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

            // Add the two series to the chart
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
            workbook.Save("SecondaryAxisPercentageFormat.xlsx");
        }
    }
}
