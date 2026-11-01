// Title: Hide the legend in an Aspose.Cells column chart to maximize the plotting area (C#)
// Common Searches: C# Aspose.Cells hide chart legend to expand plot area | remove legend from column chart using Aspose.Cells for .NET | how to disable legend in Excel chart with Aspose.Cells API | increase chart plotting space by turning off legend in Aspose.Cells | Aspose.Cells chart.ShowLegend false example in C#
// Tags: chart.ShowLegend property Aspose.Cells | disable legend column chart .NET | expand plot area Aspose.Cells chart | Aspose.Cells column chart formatting | Excel chart legend removal C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, inserts a column chart, sets its data range, disables the legend by setting chart.ShowLegend to false, and saves the result as ChartWithoutLegend.xlsx.
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

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Hide the legend to maximize the plotting area
            chart.ShowLegend = false;

            // Define output file path
            string outputPath = "ChartWithoutLegend.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
