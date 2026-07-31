// Title: C# Aspose.Cells – Hide Chart Legend When Series Count Exceeds Five
// Description: Demonstrates how to create a workbook, add a column chart with seven data series, and automatically hide the legend if the chart contains more than five series by using the ShowLegend property.
// Keywords: Aspose.Cells hide legend C# | conditional chart legend Aspose.Cells | chart series count hide legend .NET | ShowLegend property Aspose.Cells | column chart without legend Aspose
// Common Searches: Aspose.Cells hide legend for many series | C# hide chart legend based on series number | Aspose.Cells conditional legend visibility | remove legend when chart has >5 series | ShowLegend false Aspose.Cells example
// Developer Intent: Automatically suppress a chart legend when the number of series is greater than five.
// Use Cases: Generate Excel reports where charts with numerous series remain readable by omitting the legend. | Build dashboards that adapt legend visibility based on dynamic data sets. | Apply a workbook‑wide rule to hide legends on any chart exceeding a predefined series threshold.
// AI Prompts: Create C# code with Aspose.Cells that adds a line chart and disables its legend when the series count is over five. | Show how to loop through all charts in a workbook and set ShowLegend = false for charts with more than five series. | Explain the steps to check NSeries.Count and control legend visibility in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add a column chart with seven data series, and automatically hide the legend if the chart contains more than five series by using the ShowLegend property.
    public class HideLegendWhenManySeries
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data with 7 series (more than 5)
                // Category labels in column A
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["A5"].PutValue("Q4");

                // Series data in columns B to H (7 series)
                for (int col = 1; col <= 7; col++) // B=1, C=2, ... H=7
                {
                    sheet.Cells[0, col].PutValue($"Series{col}");
                    sheet.Cells[1, col].PutValue(10 * col);
                    sheet.Cells[2, col].PutValue(20 * col);
                    sheet.Cells[3, col].PutValue(30 * col);
                    sheet.Cells[4, col].PutValue(40 * col);
                }

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Add each series to the chart
                for (int col = 1; col <= 7; col++)
                {
                    // Data range for the current series (e.g., B2:B5, C2:C5, ...)
                    string dataRange = CellsHelper.CellIndexToName(1, col) + ":" + CellsHelper.CellIndexToName(4, col);
                    chart.NSeries.Add(dataRange, true);
                }

                // Set category (X‑axis) data
                chart.NSeries.CategoryData = "A2:A5";

                // Hide the legend if there are more than 5 series
                if (chart.NSeries.Count > 5)
                {
                    chart.ShowLegend = false; // Legend will not be displayed
                }

                // Save the workbook
                workbook.Save("ChartWithConditionalLegend.xlsx");
                Console.WriteLine("Workbook saved successfully.");
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
            HideLegendWhenManySeries.Run();
        }
    }
}
