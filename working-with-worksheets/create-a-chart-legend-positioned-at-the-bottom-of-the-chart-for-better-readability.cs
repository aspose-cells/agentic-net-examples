// Title: Create an Excel column chart with the legend positioned at the bottom using Aspose.Cells for .NET (C#)
// Common Searches: how to set chart legend to bottom in Aspose.Cells C# | Aspose.Cells place legend below column chart programmatically | C# generate Excel column chart with bottom legend using Aspose | Aspose.Cells chart legend position options .NET | example code for moving Excel chart legend to bottom with Aspose.Cells
// Tags: Aspose.Cells column chart legend bottom | configure legend location Aspose.Cells .NET | Excel chart legend placement C# | Aspose.Cells chart customization legend | chart legend layout options Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a workbook, adds sample data, builds a column chart, assigns series and category ranges, sets the legend position to the bottom, and saves the file as ChartWithBottomLegend.xlsx.
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

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(15);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Position the legend at the bottom of the chart
            chart.Legend.Position = LegendPositionType.Bottom;

            // Define output file path
            string outputPath = "ChartWithBottomLegend.xlsx";

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
