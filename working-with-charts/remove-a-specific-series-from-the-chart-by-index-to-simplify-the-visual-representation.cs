// Title: Remove a Chart Series by Index with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add a column chart with two data series, and delete a specific series using the zero‑based index (chart.NSeries.RemoveAt). The example saves the modified workbook, showing the remaining series count.
// Keywords: Aspose.Cells C# | .NET chart series removal | NSeries.RemoveAt | remove chart series by index | Excel chart manipulation Aspose | delete series from column chart | programmatic chart editing | Aspose.Cells chart API | C# Excel chart example
// Common Searches: How to delete a specific series from an Aspose.Cells chart | Aspose.Cells remove chart series by index C# | NSeries.RemoveAt example .NET | Remove first series from Excel column chart using Aspose | Aspose.Cells chart series delete code
// Developer Intent: Programmatically remove a chosen data series from an Excel chart using Aspose.Cells for .NET.
// Use Cases: Trim automatically generated charts by discarding placeholder or unwanted series before publishing. | Implement a UI option that lets users hide a series, then remove it from the workbook via code. | Prepare a workbook for PDF or image export by cleaning up extra series that could clutter the visual.
// AI Prompts: Show C# code to remove the second series from a line chart with Aspose.Cells. | Generate an Aspose.Cells snippet that removes a chart series based on its name. | Explain how to check the number of remaining series after calling NSeries.RemoveAt.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add a column chart with two data series, and delete a specific series using the zero‑based index (chart.NSeries.RemoveAt). The example saves the modified workbook, showing the remaining series count.
    public class RemoveSeriesByIndexDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("Q1");
                worksheet.Cells["A3"].PutValue("Q2");
                worksheet.Cells["A4"].PutValue("Q3");

                worksheet.Cells["B1"].PutValue("Series1");
                worksheet.Cells["B2"].PutValue(100);
                worksheet.Cells["B3"].PutValue(200);
                worksheet.Cells["B4"].PutValue(300);

                worksheet.Cells["C1"].PutValue("Series2");
                worksheet.Cells["C2"].PutValue(150);
                worksheet.Cells["C3"].PutValue(250);
                worksheet.Cells["C4"].PutValue(350);

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = worksheet.Charts[chartIndex];

                // Add two series to the chart
                chart.NSeries.Add("B2:B4", true); // Series1
                chart.NSeries.Add("C2:C4", true); // Series2
                chart.NSeries.CategoryData = "A2:A4";

                // Remove the first series (index 0)
                chart.NSeries.RemoveAt(0);

                // Display remaining series count for verification
                Console.WriteLine($"Remaining series count: {chart.NSeries.Count}");

                // Save the workbook with the modified chart
                string outputPath = "RemoveSeriesByIndexDemo.xlsx";
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
            RemoveSeriesByIndexDemo.Run();
        }
    }
}
