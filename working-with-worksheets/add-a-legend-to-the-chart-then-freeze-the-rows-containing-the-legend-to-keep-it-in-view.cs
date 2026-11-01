// Title: Add a right‑side legend to a column chart and freeze the rows containing the legend using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that creates a column chart, shows a legend on the right side, and freezes the top rows so the legend stays visible while scrolling. | Write a program that adds a legend to an Excel chart and applies FreezePanes to keep the legend in view using Aspose.Cells for .NET.
// Common Searches: asp.net add legend to column chart and freeze top rows using Aspose.Cells | C# Aspose.Cells keep chart legend visible while scrolling | how to freeze panes after inserting a chart legend with Aspose.Cells | Aspose.Cells set legend position right and freeze rows example
// Tags: column chart legend Aspose.Cells C# | freeze panes rows Aspose.Cells | legend position right Aspose.Cells | keep chart legend visible scrolling Aspose.Cells | add legend to chart Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, fills it with sample data, inserts a column chart, enables a right‑side legend, freezes the first seven rows so the legend remains visible during scrolling, and saves the file as ChartWithLegend.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart (rows 5‑20, columns 0‑7)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the series
            chart.NSeries.Add("B2:B4", true);
            // Category data is inferred from the first column (A2:A4) by default,
            // so explicit assignment is not required.

            // Enable the legend and position it on the right side of the chart
            chart.ShowLegend = true;
            chart.Legend.Position = LegendPositionType.Right;

            // Freeze the top 7 rows so the legend (if placed there) stays visible while scrolling
            sheet.FreezePanes(7, 0, 0, 0);

            // Ensure the output directory exists
            string outputPath = "ChartWithLegend.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
