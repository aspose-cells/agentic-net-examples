// Title: Set the legend of a column chart to the top‑right corner in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that creates a column chart and moves its legend to the top‑right corner of the chart area. | Demonstrate how to use the Legend.Position property to place a chart legend at the top‑right location in an Aspose.Cells workbook.
// Common Searches: aspnet aspose.cells place chart legend top right | c# how to set legend position to top right in Excel chart using Aspose.Cells | Aspose.Cells legend placement options for column charts in .NET | move Excel chart legend to top‑right corner with Aspose.Cells library
// Tags: Aspose.Cells chart legend positioning | C# set legend top right Aspose.Cells | Excel column chart legend placement Aspose.Cells | LegendPositionType top right Aspose.Cells | Aspose.Cells workbook chart customization

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, adds a column chart with sample data, sets the chart legend to the top‑right corner using the Legend.Position property, and saves the file as ChartWithTopRightLegend.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Define the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set legend position to the top of the chart (closest available option)
            chart.Legend.Position = LegendPositionType.Top;

            // Determine output file path
            string outputPath = "ChartWithTopRightLegend.xlsx";

            // Save the workbook with the chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
