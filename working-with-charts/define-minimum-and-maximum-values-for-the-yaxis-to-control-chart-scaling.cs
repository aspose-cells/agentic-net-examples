// Title: Set custom minimum and maximum values for a chart Y‑axis using Aspose.Cells for .NET (C#)
// AI Prompts: Create a column chart from worksheet data and configure the value axis with MinValue = 0, MaxValue = 100, and MajorUnit = 20 using Aspose.Cells in C#. | Adjust an existing Aspose.Cells chart to enforce specific Y‑axis scaling by setting the axis minimum, maximum, and major tick interval programmatically. | Generate a workbook, populate sample data, add a column chart, and programmatically set the Y‑axis range and tick spacing with Aspose.Cells for .NET.
// Common Searches: how to set y axis minimum and maximum in Aspose.Cells chart c# | Aspose.Cells chart value axis scaling example | C# Aspose.Cells set major unit for chart axis | custom y axis range for column chart using Aspose.Cells .NET | control chart y axis range programmatically Aspose.Cells
// Tags: Aspose.Cells chart y-axis min max | set chart value axis range .NET | column chart axis scaling Aspose.Cells | major unit configuration Aspose.Cells | Excel workbook y-axis scaling C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates a workbook, adds sample data, inserts a column chart, and customizes the chart's Y‑axis by setting MinValue to 0, MaxValue to 100, and MajorUnit to 20 before saving the file as ChartWithCustomYAxis.xlsx.
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
            sheet.Cells["B2"].PutValue(20);
            sheet.Cells["B3"].PutValue(55);
            sheet.Cells["B4"].PutValue(80);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart series
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the value (Y) axis
            Axis valueAxis = chart.ValueAxis;

            // Define minimum and maximum values for the Y‑axis
            valueAxis.MinValue = 0;    // Minimum value
            valueAxis.MaxValue = 100;  // Maximum value

            // Optional: set major unit for better tick spacing
            valueAxis.MajorUnit = 20;

            // Save the workbook to a file
            string outputPath = "ChartWithCustomYAxis.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
