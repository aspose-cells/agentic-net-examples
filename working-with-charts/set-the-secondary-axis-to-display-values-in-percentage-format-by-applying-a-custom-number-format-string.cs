// Title: Format Secondary Axis as Percentage in an Aspose.Cells Column Chart (C#)
// Description: Demonstrates how to apply a custom number format (e.g., "0.00%") to the secondary value axis of a column chart using Aspose.Cells for .NET, including data setup, secondary series plotting, axis title, and workbook saving.
// Keywords: Aspose.Cells | C# chart formatting | secondary axis percentage | custom number format | TickLabels.NumberFormat | PlotOnSecondAxis | column chart Aspose.Cells | .NET Excel API | Excel secondary value axis | percentage tick labels
// Common Searches: Aspose.Cells set secondary axis to percent | C# chart secondary axis number format Aspose.Cells | how to display secondary axis values as % in Aspose.Cells | custom number format for secondary value axis .NET | Aspose.Cells column chart secondary axis formatting
// Developer Intent: Apply a percentage number format to the tick labels of a chart's secondary axis in Aspose.Cells for .NET.
// Use Cases: Financial report showing revenue on the primary axis and profit margin on the secondary axis, with margins displayed as percentages. | Sales dashboard where units sold are plotted on the primary axis and conversion rates on the secondary axis, formatted as % values. | Performance analysis chart that compares raw scores with growth ratios, using a secondary axis to present the ratios in percent format.
// AI Prompts: Show C# code to set a custom "0.00%" format for secondary axis tick labels in an Aspose.Cells column chart. | Explain how to plot a series on the secondary value axis and format it as a percentage using Aspose.Cells for .NET. | Provide steps to apply different number formats to primary and secondary axes in an Aspose.Cells chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to apply a custom number format (e.g., "0.00%") to the secondary value axis of a column chart using Aspose.Cells for .NET, including data setup, secondary series plotting, axis title, and workbook saving.
    public class SecondaryAxisPercentageFormatDemo
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

                // Add the two series to the chart
                chart.NSeries.Add("B2:B4", true); // primary series
                chart.NSeries.Add("C2:C4", true); // secondary series
                chart.NSeries.CategoryData = "A2:A4";

                // Plot the second series on the secondary value axis
                chart.NSeries[1].PlotOnSecondAxis = true;

                // Apply a custom number format to the secondary value axis tick labels
                // This will display the axis values as percentages (e.g., 25% instead of 0.25)
                chart.SecondValueAxis.TickLabels.NumberFormat = "0.00%";

                // Optionally, give the secondary axis a title
                chart.SecondValueAxis.Title.Text = "Secondary Axis (Percentage)";

                // Save the workbook
                workbook.Save("SecondaryAxisPercentageFormatDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
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
            SecondaryAxisPercentageFormatDemo.Run();
        }
    }
}
