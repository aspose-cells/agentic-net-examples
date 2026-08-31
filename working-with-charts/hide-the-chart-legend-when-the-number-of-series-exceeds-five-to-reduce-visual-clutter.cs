// Title: Hide a chart legend in Aspose.Cells when more than five series are present (C#)
// AI Prompts: Generate C# code using Aspose.Cells that creates a column chart and automatically disables the legend if the chart contains more than five data series. | Demonstrate how to check the NSeries.Count of an Aspose.Cells chart and set the ShowLegend property conditionally in a .NET application.
// Common Searches: Aspose.Cells C# hide chart legend if series count exceeds five | How to conditionally display Excel chart legend with Aspose.Cells .NET | Set ShowLegend based on number of series in Aspose.Cells chart | C# Aspose.Cells column chart legend auto hide for many series
// Tags: conditional legend visibility Aspose.Cells | chart legend hide based on series count | Aspose.Cells NSeries count check | C# dynamic Excel chart legend | column chart auto hide legend Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a workbook with six data series, adds a column chart, and uses the NSeries.Count property to set ShowLegend to false when more than five series exist, then saves the file.
    public class ConditionalLegendDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data with six series (more than five)
                // Category column
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["A5"].PutValue("Q4");

                // Six series columns (B to G)
                string[] seriesHeaders = { "Series1", "Series2", "Series3", "Series4", "Series5", "Series6" };
                Random rnd = new Random();
                for (int col = 1; col <= seriesHeaders.Length; col++)
                {
                    // Header
                    sheet.Cells[0, col].PutValue(seriesHeaders[col - 1]);
                    // Random values for each category
                    for (int row = 1; row <= 4; row++)
                    {
                        sheet.Cells[row, col].PutValue(rnd.Next(10, 100));
                    }
                }

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the category (X) data
                chart.NSeries.CategoryData = "A2:A5";

                // Add each series to the chart
                for (int col = 1; col <= seriesHeaders.Length; col++)
                {
                    // Data range for the current series (e.g., B2:B5, C2:C5, ...)
                    string dataRange = CellsHelper.CellIndexToName(0, col) + "2:" + CellsHelper.CellIndexToName(0, col) + "5";
                    chart.NSeries.Add(dataRange, true);
                }

                // Hide the legend if the number of series exceeds five
                chart.ShowLegend = chart.NSeries.Count <= 5;

                // Save the workbook
                string outputPath = "ConditionalChartLegend.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
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
            ConditionalLegendDemo.Run();
        }
    }
}
