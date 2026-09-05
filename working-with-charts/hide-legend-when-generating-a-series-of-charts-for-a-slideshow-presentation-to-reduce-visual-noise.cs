// Title: Hide chart legends while generating multiple column charts for a slideshow with Aspose.Cells for .NET
// AI Prompts: Create three column charts in a worksheet and set ShowLegend = false for each chart using the Aspose.Cells C# API. | Write a C# loop that adds column charts, assigns data and category ranges, and disables the legend to reduce visual noise. | Generate a workbook with sample data and export it as an Excel file where all chart legends are hidden for presentation purposes.
// Common Searches: asp.net hide legend for each chart using Aspose.Cells | c# Aspose.Cells generate multiple column charts without legends | how to disable chart legend in Aspose.Cells loop for slideshow | Aspose.Cells ShowLegend false example for presentation charts | create series of charts with hidden legends in Excel using Aspose.Cells C#
// Tags: Aspose.Cells hide chart legend | Aspose.Cells create column chart | Aspose.Cells multiple charts loop | Aspose.Cells slideshow chart formatting | Aspose.Cells ShowLegend property

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSlideshow
{
    // Demonstrates how to build a workbook, populate sample data, add three column charts positioned sequentially, assign data ranges, and suppress each chart’s legend by setting ShowLegend to false, then save the file as SlideshowCharts.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data for the charts
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

            // Create multiple charts (e.g., three charts) and hide their legends
            for (int i = 0; i < 3; i++)
            {
                // Add a column chart; position varies per iteration
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5 + i * 15, 0, 15 + i * 15, 5);
                Chart chart = sheet.Charts[chartIndex];

                // Set data range for the chart
                chart.NSeries.Add("B2:C4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";     // Categories

                // Hide the legend to reduce visual noise
                chart.ShowLegend = false;
            }

            // Save the workbook containing the charts
            workbook.Save("SlideshowCharts.xlsx");
        }
    }
}
