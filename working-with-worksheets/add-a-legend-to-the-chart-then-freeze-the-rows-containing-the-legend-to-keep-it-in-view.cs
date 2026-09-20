// Title: How to add a bottom legend to a column chart and freeze the legend rows in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a column chart, positions its legend at the bottom, and applies FreezePanes to keep the legend rows visible while scrolling, using Aspose.Cells. | Write a .NET snippet that adds a bottom‑positioned legend to a chart and then freezes the first five rows of the worksheet so the legend stays in view, leveraging Aspose.Cells APIs.
// Common Searches: asp.net add bottom legend to column chart and freeze rows with Aspose.Cells | how to keep chart legend visible while scrolling using Aspose.Cells FreezePanes | set legend position bottom and freeze top rows in Excel file Aspose.Cells C# | Aspose.Cells example freezing rows above a chart legend | C# Aspose.Cells column chart with legend and frozen header rows
// Tags: Aspose.Cells column chart legend bottom | Aspose.Cells FreezePanes rows | C# chart legend positioning Aspose.Cells | freeze worksheet header rows Aspose.Cells | Excel workbook chart legend visibility .NET

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// Creates a workbook, inserts sample data, adds a column chart with a bottom legend, freezes the first five rows so the legend remains visible during scrolling, and saves the file as ChartWithLegend.xlsx using Aspose.Cells for .NET.
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

            // Add a column chart to the worksheet (rows 5‑20, columns A‑G)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";     // Categories

            // Enable and position the legend
            chart.ShowLegend = true;
            // Use LegendPositionType for compatibility with older Aspose.Cells versions
            chart.Legend.Position = LegendPositionType.Bottom;

            // Freeze rows above the chart (rows 0‑4) so the chart stays visible while scrolling
            int totalRows = sheet.Cells.MaxDataRow + 1;
            int totalColumns = sheet.Cells.MaxDataColumn + 1;
            sheet.FreezePanes(5, 0, totalRows, totalColumns);

            // Save the workbook
            workbook.Save("ChartWithLegend.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
