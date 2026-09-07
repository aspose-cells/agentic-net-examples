// Title: Remove a chart series by index from an Excel column chart using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that builds a column chart with two data series using Aspose.Cells, then deletes the series at index 0 with NSeries.RemoveAt and saves the workbook. | Show how to call the NSeries.RemoveAt method on an Aspose.Cells chart to eliminate a specific series and output the remaining series count. | Demonstrate removing a chart series by its zero‑based index in Aspose.Cells and verify the chart updates correctly.
// Common Searches: Aspose.Cells C# remove first series from column chart | How to delete a specific series from an Excel chart using Aspose.Cells .NET | Aspose.Cells NSeries.RemoveAt example C# | Remove chart series by index in Aspose.Cells workbook | C# Aspose.Cells chart series count after removal
// Tags: Aspose.Cells chart series removal | NSeries.RemoveAt method C# | column chart series manipulation Aspose.Cells | Excel chart series deletion .NET | Aspose.Cells workbook chart editing

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a workbook, adds category data and two series, builds a column chart, removes the first series using NSeries.RemoveAt(0), prints the remaining series count, and saves the file as RemoveSeriesByIndexDemo.xlsx.
    public class RemoveSeriesByIndexDemo
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            try
            {
                Run();
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

            // Populate sample data for two series
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

            // Set data ranges for the two series
            chart.NSeries.Add("B2:B4", true); // Series1
            chart.NSeries.Add("C2:C4", true); // Series2
            chart.NSeries.CategoryData = "A2:A4";

            // Remove the first series (index 0) to simplify the chart
            int seriesToRemove = 0;
            chart.NSeries.RemoveAt(seriesToRemove);

            // Optional: display remaining series count
            Console.WriteLine($"Remaining series count: {chart.NSeries.Count}");

            // Save the workbook
            string outputPath = "RemoveSeriesByIndexDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
