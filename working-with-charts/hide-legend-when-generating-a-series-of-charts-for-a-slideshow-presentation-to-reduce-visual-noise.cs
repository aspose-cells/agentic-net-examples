// Title: Hide Legends in Multiple Column Charts Using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a new workbook, add sample data, generate three column charts each bound to a separate series, position them programmatically, disable their legends with ShowLegend = false, and save the result as MultipleCharts_NoLegend.xlsx.
// Keywords: Aspose.Cells C# chart legend | hide chart legend Aspose.Cells | multiple column charts .NET | programmatic chart positioning Excel | Aspose.Cells chart automation | Excel workbook without legend | C# generate charts Aspose.Cells | presentation ready charts Aspose
// Common Searches: Aspose.Cells hide legend multiple charts | C# create several column charts in one sheet | set ShowLegend false Aspose.Cells | position charts programmatically Aspose.Cells | generate Excel charts for slideshow
// Developer Intent: Create several column charts in a single worksheet and suppress their legends to keep the visual layout clean for presentation purposes.
// Use Cases: Building a slide‑deck workbook where each slide contains a chart focused on a single metric, without a legend that would distract the audience. | Automating a reporting tool that outputs multiple metric charts on one sheet, where the series name is already evident from the chart title. | Designing a dashboard that places several charts side‑by‑side and removes legends to maximize data‑area space.
// AI Prompts: Show C# code that adds three column charts with Aspose.Cells and hides each chart's legend. | Explain how to use ShowLegend = false inside a loop when creating charts from different data series. | Provide guidance on positioning multiple charts on a worksheet and removing legends for a clean presentation layout.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a new workbook, add sample data, generate three column charts each bound to a separate series, position them programmatically, disable their legends with ShowLegend = false, and save the result as MultipleCharts_NoLegend.xlsx.
    public class HideLegendInMultipleCharts
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully: MultipleCharts_NoLegend.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the charts
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);
            sheet.Cells["D1"].PutValue("Series3");
            sheet.Cells["D2"].PutValue(12);
            sheet.Cells["D3"].PutValue(22);
            sheet.Cells["D4"].PutValue(32);

            // Define chart positions (row, column, row2, column2) for three charts
            int[,] chartPositions = new int[,]
            {
                {5, 0, 15, 5},
                {5, 6, 15, 11},
                {5, 12, 15, 17}
            };

            // Create three column charts, each using a different series, and hide their legends
            for (int i = 0; i < chartPositions.GetLength(0); i++)
            {
                int row = chartPositions[i, 0];
                int col = chartPositions[i, 1];
                int row2 = chartPositions[i, 2];
                int col2 = chartPositions[i, 3];

                // Add a chart of type Column
                int chartIndex = sheet.Charts.Add(ChartType.Column, row, col, row2, col2);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the current series (Series columns are B, C, D respectively)
                char seriesColumn = (char)('B' + i);
                string seriesRange = $"{seriesColumn}2:{seriesColumn}4";
                chart.NSeries.Add(seriesRange, true);
                chart.NSeries.CategoryData = "A2:A4";

                // Hide the legend to reduce visual noise
                chart.ShowLegend = false;
            }

            // Save the workbook containing the charts
            workbook.Save("MultipleCharts_NoLegend.xlsx");
        }
    }
}
