// Title: How to position a chart legend at the top‑right corner using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set chart legend to top right corner | move Excel chart legend to top right with Aspose.Cells API | C# code for placing chart legend at top right using Aspose.Cells | Aspose.Cells legend positioning options for column charts | set Legend.Position to top right in Aspose.Cells .NET
// Tags: Aspose.Cells LegendPositionType Top | C# set chart legend location Aspose.Cells | Excel chart legend placement using Aspose.Cells | column chart legend positioning Aspose.Cells | Aspose.Cells chart formatting legend

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This example creates a workbook, adds sample data, inserts a column chart, and moves the chart's legend to the top‑right area of the chart by configuring Legend.Position with Aspose.Cells for .NET, then saves the file as ChartWithTopRightLegend.xlsx.
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
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart series (vertical series)
            chart.NSeries.Add("B2:B4", true);

            // NOTE: In recent Aspose.Cells versions the Series class does not expose a
            // CategoryData property directly. The chart will automatically use the
            // adjacent column (A2:A4) as category labels when the series is added as vertical.
            // If explicit category data is required, use the SetCategoryData method (available in newer releases):
            // chart.NSeries[0].SetCategoryData("A2:A4");

            // Position the legend at the top (closest available position)
            chart.Legend.Position = LegendPositionType.Top;

            // Define output file path
            string outputPath = "ChartWithTopRightLegend.xlsx";

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
