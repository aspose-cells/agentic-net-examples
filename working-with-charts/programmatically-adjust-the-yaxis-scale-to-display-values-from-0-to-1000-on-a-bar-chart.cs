// Title: How to set a fixed Y‑axis range of 0‑1000 and define a major unit for a column chart using Aspose.Cells for .NET (C#)
// AI Prompts: Create a column chart from worksheet data and set the value axis minimum to 0 and maximum to 1000 using Aspose.Cells in C#. | Configure the Y‑axis major unit to 200 for a bar chart generated with Aspose.Cells in a .NET project. | Produce an XLSX file containing a bar chart that uses a fixed Y‑axis scale of 0‑1000 via Aspose.Cells for C#.
// Common Searches: Aspose.Cells C# how to set chart Y axis minimum and maximum values | custom Y axis range for column chart using Aspose.Cells .NET | C# example for setting major unit on Aspose.Cells chart axis | programmatically adjust Y‑axis scale of bar chart in Aspose.Cells | fixed Y axis limits for Excel chart with Aspose.Cells API
// Tags: Aspose.Cells chart value axis range | column chart Y-axis scaling Aspose.Cells | chart major unit configuration Aspose.Cells | C# generate Excel bar chart with fixed Y-axis | Aspose.Cells axis formatting example

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new workbook, adds sample data, generates a column (bar) chart, sets the chart's value axis minimum to 0, maximum to 1000, defines a major unit of 200, and saves the workbook as BarChart_With_CustomYAxis.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Item 1");
            sheet.Cells["A3"].PutValue("Item 2");
            sheet.Cells["A4"].PutValue("Item 3");
            sheet.Cells["B2"].PutValue(200);
            sheet.Cells["B3"].PutValue(600);
            sheet.Cells["B4"].PutValue(900);

            // Add a column (bar) chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set Y‑axis (value axis) scale from 0 to 1000
            chart.ValueAxis.MinValue = 0;
            chart.ValueAxis.MaxValue = 1000;

            // Optional: set major unit for better readability
            chart.ValueAxis.MajorUnit = 200;

            // Save the workbook to a file
            workbook.Save("BarChart_With_CustomYAxis.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
