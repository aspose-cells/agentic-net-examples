// Title: How to filter out categories containing "Total" from a column chart using Aspose.Cells in C#
// AI Prompts: Write C# code that uses Aspose.Cells to add a CategoryFilter to a column chart so that any category label containing the word "Total" is excluded from the chart. | Update the given Aspose.Cells example to hide data points whose category cells include "Total" by applying the chart's CategoryFilter method.
// Common Searches: Aspose.Cells C# filter chart categories containing specific text | Exclude "Total" labels from an Excel column chart with Aspose.Cells | CategoryFilter method usage for Aspose.Cells charts in .NET | How to hide rows with "Total" in chart series using Aspose.Cells C#
// Tags: Aspose.Cells CategoryFilter column chart | C# filter chart categories by label text | exclude Total categories Aspose.Cells | Excel chart series filtering with Aspose.Cells | apply CategoryFilter method C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, fills columns A and B with category and value data, adds a column chart for the range A2:B6, and demonstrates how to apply Aspose.Cells' CategoryFilter to remove any categories whose labels contain the word "Total" before saving the file as FilteredChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];
            var cells = sheet.Cells;

            // Populate header
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value");

            // Sample data
            string[] categories = { "Jan", "Feb", "Total Q1", "Mar", "Total Q2" };
            double[] values = { 10, 20, 30, 40, 50 };

            // Fill data into the worksheet
            for (int i = 0; i < categories.Length; i++)
            {
                cells[i + 1, 0].PutValue(categories[i]); // Column A
                cells[i + 1, 1].PutValue(values[i]);    // Column B
            }

            // Add a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            var chart = sheet.Charts[chartIdx];

            // Use a range that includes both categories (A) and values (B)
            chart.NSeries.Add("A2:B6", true);

            // Save the workbook
            string outputPath = "FilteredChart.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
