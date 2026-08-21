// Title: Hide chart legend in Aspose.Cells for .NET to maximize plot area
// Description: Demonstrates how to create a workbook, add sample data, insert a column chart, bind the data range, disable the legend using the ShowLegend property, and save the result as ChartWithoutLegend.xlsx.
// Keywords: Aspose.Cells hide legend | Aspose.Cells ShowLegend false | remove chart legend C# | Aspose.Cells chart area optimization | Aspose.Cells .NET chart formatting
// Common Searches: How to hide a legend in an Aspose.Cells chart (C#) | Aspose.Cells remove chart legend to increase plot area | Set ShowLegend = false in Aspose.Cells | Maximize Excel chart area by disabling legend with Aspose.Cells | C# Aspose.Cells chart without legend
// Developer Intent: Disable the chart legend so the plot area occupies the full chart space.
// Use Cases: Compact dashboards where legends are redundant | Space‑efficient Excel reports with multiple charts | Print‑ready charts that avoid unnecessary legend space | Embedding charts in PDFs where legends would be clipped
// AI Prompts: Generate C# code with Aspose.Cells to create a line chart and hide its legend. | Explain the ShowLegend property and list other chart elements that can be toggled in Aspose.Cells. | Show how to automatically resize a chart after hiding the legend to fill the empty area.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook, add sample data, insert a column chart, bind the data range, disable the legend using the ShowLegend property, and save the result as ChartWithoutLegend.xlsx.
public class RemoveLegendDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Hide the legend completely to maximize the plotting area
            chart.ShowLegend = false;

            // Save the workbook with the chart that has no legend
            workbook.Save("ChartWithoutLegend.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        RemoveLegendDemo.Run();
    }
}
