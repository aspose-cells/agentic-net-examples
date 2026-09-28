// Title: Add a column chart from a cell range and set its legend to the right side using Aspose.Cells for .NET
// AI Prompts: Create a column chart from the range B2:B4 on the first worksheet and move the legend to the right using Aspose.Cells C#. | Generate a chart, bind it to a data range, and customize the legend position to the right side in a .NET workbook with Aspose.Cells.
// Common Searches: Aspose.Cells how to bind column chart to specific cells in C# | set legend position to right for chart in Aspose.Cells .NET | create chart from data range and customize legend location using Aspose.Cells for .NET | add column chart to worksheet programmatically with Aspose.Cells C# example
// Tags: Aspose.Cells create column chart from range | Aspose.Cells set chart legend position | Aspose.Cells chart NSeries binding | Aspose.Cells .xlsx chart generation | Aspose.Cells customize chart legend

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// The sample creates a new workbook, fills month‑sales data, adds a column chart covering rows 5‑20, binds the series to cells B2:B4, positions the legend on the right side, and saves the file as ChartWithCustomLegend.xlsx.
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
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(130);

            // Add a column chart (positioned from row 5, column 0 to row 20, column 7)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series (values). Categories are taken from the first column by default.
            chart.NSeries.Add("B2:B4", true);

            // Customize the legend position (place it on the right side of the chart)
            chart.Legend.Position = LegendPositionType.Right;

            // Save the workbook to a file
            workbook.Save("ChartWithCustomLegend.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
