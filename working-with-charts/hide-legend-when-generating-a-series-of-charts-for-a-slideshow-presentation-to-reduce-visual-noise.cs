// Title: Hide Chart Legends in Multiple Aspose.Cells Charts (C#) for Slide Presentations
// Description: Demonstrates how to create a workbook, add column, line, and pie charts, and programmatically hide each chart's legend by setting the ShowLegend property to false. The example uses a loop to apply the setting to several charts and saves the result as MultipleCharts_NoLegend.xlsx.
// Keywords: Aspose.Cells chart legend | hide legend C# | ShowLegend false | multiple charts .NET | Excel chart formatting Aspose | C# Aspose.Cells example | slide deck charts | PowerPoint export without legend | Aspose.Cells API chart settings
// Common Searches: Aspose.Cells hide legend C# | remove legend from all charts Aspose.Cells .NET | ShowLegend property example | create multiple charts without legend Aspose | C# generate Excel charts for PowerPoint without legend
// Developer Intent: Generate several charts in an Excel workbook and suppress their legends to keep slide visuals clean and uncluttered.
// Use Cases: Building a slide deck where chart legends would overlap slide content. | Producing a compact Excel report that displays only data series without extra labels. | Exporting workbook charts to PowerPoint or PDF while maintaining a minimalist design.
// AI Prompts: Provide C# code that creates column, line, and pie charts with Aspose.Cells and disables their legends. | How can I hide legends for all charts in an Aspose.Cells workbook using a loop? | Show an example of conditionally hiding legends only for specific chart types in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add column, line, and pie charts, and programmatically hide each chart's legend by setting the ShowLegend property to false. The example uses a loop to apply the setting to several charts and saves the result as MultipleCharts_NoLegend.xlsx.
    public class HideLegendInMultipleCharts
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the charts
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["B1"].PutValue("Series 1");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["C1"].PutValue("Series 2");
                sheet.Cells["C2"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["C4"].PutValue(35);

                // Define chart types to create
                ChartType[] chartTypes = new ChartType[]
                {
                    ChartType.Column,
                    ChartType.Line,
                    ChartType.Pie
                };

                // Create each chart with its own position
                for (int i = 0; i < chartTypes.Length; i++)
                {
                    // Add a chart to the worksheet
                    int chartIndex = sheet.Charts.Add(chartTypes[i], 5 + i * 15, 0, 20 + i * 15, 10);
                    Chart chart = sheet.Charts[chartIndex];

                    // Set data range for the chart
                    if (chartTypes[i] == ChartType.Pie)
                    {
                        chart.NSeries.Add("B2:B4", true);
                        chart.NSeries.CategoryData = "A2:A4";
                    }
                    else
                    {
                        chart.NSeries.Add("B2:B4", true);
                        chart.NSeries[0].Name = "Series 1";
                        chart.NSeries.Add("C2:C4", true);
                        chart.NSeries[1].Name = "Series 2";
                        chart.NSeries.CategoryData = "A2:A4";
                    }

                    // Hide the legend
                    chart.ShowLegend = false;
                }

                // Save the workbook containing all charts
                workbook.Save("MultipleCharts_NoLegend.xlsx");
                Console.WriteLine("Workbook saved successfully as MultipleCharts_NoLegend.xlsx");
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
