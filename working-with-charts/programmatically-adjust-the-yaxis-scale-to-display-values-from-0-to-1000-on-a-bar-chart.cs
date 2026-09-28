// Title: Set a fixed Y‑axis range of 0 to 1000 on a column chart using Aspose.Cells for .NET (C#)
// AI Prompts: Create an Excel workbook, add a column chart, and programmatically set the value axis minimum to 0 and maximum to 1000 with Aspose.Cells in C#. | Disable automatic scaling and define a major unit of 100 for the Y‑axis of a bar chart using Aspose.Cells for .NET. | Generate BarChart_With_CustomYAxis.xlsx containing sample data and a column chart whose Y‑axis is fixed from 0 to 1000.
// Common Searches: aspnet set y axis minimum and maximum for chart using Aspose.Cells | c# Aspose.Cells column chart custom value axis range | how to disable automatic axis scaling in Aspose.Cells chart | Aspose.Cells set major unit for chart Y axis programmatically
// Tags: Aspose.Cells set value axis range | column chart custom Y axis Aspose.Cells | disable automatic axis scaling Aspose.Cells | set chart major unit C# Aspose.Cells | export chart to Excel with fixed axis Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, fills it with sample data, adds a column chart, and configures the chart's value axis to a fixed range of 0‑1000 with a major unit of 100, then saves the file as BarChart_With_CustomYAxis.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the bar chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(200);
            sheet.Cells["B3"].PutValue(600);
            sheet.Cells["B4"].PutValue(900);

            // Add a column (vertical bar) chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Adjust the Y‑axis (value axis) scale to display from 0 to 1000
            Axis valueAxis = chart.ValueAxis;
            valueAxis.IsAutomaticMinValue = false;   // Disable automatic minimum
            valueAxis.IsAutomaticMaxValue = false;   // Disable automatic maximum
            valueAxis.MinValue = 0;                   // Set minimum value
            valueAxis.MaxValue = 1000;                // Set maximum value

            // Optional: set major unit for better tick marks
            valueAxis.IsAutomaticMajorUnit = false;
            valueAxis.MajorUnit = 100;

            // Save the workbook to a file
            workbook.Save("BarChart_With_CustomYAxis.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
