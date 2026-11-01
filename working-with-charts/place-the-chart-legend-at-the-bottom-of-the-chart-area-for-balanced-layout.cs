// Title: How to position a column chart legend at the bottom using Aspose.Cells for .NET (C#)
// AI Prompts: Create an Excel workbook with sample data, add a column chart, and set its legend to the bottom using Aspose.Cells in C#. | Write C# code that generates a workbook, inserts a column chart, and moves the chart legend to the bottom of the chart area with the Aspose.Cells API. | Produce a file named ChartWithBottomLegend.xlsx where the column chart's legend is positioned at the bottom, leveraging Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# set chart legend position to bottom | programmatically move Excel chart legend to bottom using Aspose.Cells | column chart legend placement bottom Aspose.Cells .NET example | how to change legend position in Aspose.Cells chart | C# Aspose.Cells place chart legend below chart area
// Tags: Aspose.Cells set chart legend position | C# column chart legend bottom | Aspose.Cells chart legend placement | Excel chart legend positioning .NET | Aspose.Cells generate column chart with bottom legend

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates a workbook, adds sample data, inserts a column chart, positions its legend at the bottom of the chart area, and saves the result as ChartWithBottomLegend.xlsx using Aspose.Cells for .NET.
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

            // Add sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(15);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Place the legend at the bottom of the chart area
            chart.Legend.Position = Aspose.Cells.Charts.LegendPositionType.Bottom;

            // Save the workbook to a file
            workbook.Save("ChartWithBottomLegend.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
