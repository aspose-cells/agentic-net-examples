// Title: How to bind a column chart to the cell range A1:B12 with Chart.SetChartDataRange in Aspose.Cells for .NET
// Common Searches: c# Aspose.Cells set chart data range to A1:B12 | using Chart.SetChartDataRange to bind column chart in .NET | example of vertical series orientation with SetChartDataRange Aspose.Cells | programmatically assign worksheet range as chart source in C# | Aspose.Cells chart data source binding tutorial for column charts
// Tags: Chart.SetChartDataRange Aspose.Cells .NET | bind column chart to worksheet range | chart data source A1:B12 example | vertical series orientation chart API | create column chart programmatically Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new workbook, adds a column chart to the first worksheet, and uses Chart.SetChartDataRange with a vertical series flag to bind the chart's data series to cells A1:B12, then saves the file as ChartDataRange.xlsx.
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

            // Add a column chart (you can choose any chart type)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Bind the chart to the data range A1:B12.
            // The second argument specifies whether the data series are arranged vertically (in columns).
            chart.SetChartDataRange("A1:B12", true);

            // Save the workbook
            string outputPath = "ChartDataRange.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
