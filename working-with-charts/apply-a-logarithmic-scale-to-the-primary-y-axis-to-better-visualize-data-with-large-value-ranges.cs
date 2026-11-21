// Title: Apply a logarithmic scale to the primary Y‑axis of a column chart using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set primary value axis to logarithmic scale | how to enable log scale on a column chart axis with Aspose.Cells | example code for logarithmic Y axis in Excel chart using Aspose.Cells .NET | create column chart with large value range and log axis in C# Aspose.Cells | Aspose.Cells chart axis IsLogarithmic property usage example
// Tags: logarithmic scaling for chart value axis Aspose.Cells | primary value axis log scaling .NET | column chart axis formatting C# Aspose.Cells | Excel workbook create column chart with log Y axis | Aspose.Cells chart axis IsLogarithmic property | logarithmic Y axis for large data range Excel .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new workbook, adds sample data, inserts a column chart, binds the series, enables a logarithmic scale on the chart's primary Y (value) axis by setting IsLogarithmic to true, and saves the file as LogarithmicScaleChart.xlsx.
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

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Low");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Medium");
            sheet.Cells["B3"].PutValue(1000);
            sheet.Cells["A4"].PutValue("High");
            sheet.Cells["B4"].PutValue(100000);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Set data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Apply logarithmic scale to the primary value axis
            Axis valueAxis = chart.ValueAxis;
            valueAxis.IsLogarithmic = true; // enable logarithmic scale

            // Save the workbook
            string outputPath = "LogarithmicScaleChart.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
