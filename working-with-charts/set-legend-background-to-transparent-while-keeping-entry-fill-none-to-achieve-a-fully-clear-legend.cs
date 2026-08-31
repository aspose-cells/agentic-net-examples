// Title: How to make a chart legend completely transparent in Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that creates a column chart and sets the legend background to transparent. | Demonstrate iterating over LegendEntry objects to remove both background and text fill in an Aspose.Cells chart. | Provide a full example that saves the workbook as an .xlsx file with a fully clear legend.
// Common Searches: Aspose.Cells C# set chart legend background mode to transparent | remove fill from legend entries in Aspose.Cells chart using C# | transparent legend for Excel chart generated with Aspose.Cells .NET | how to make legend entries have no text fill in Aspose.Cells | save workbook with clear legend using Aspose.Cells C#
// Tags: chart legend background mode transparent Aspose.Cells | legend entry no fill Aspose.Cells C# | Aspose.Cells set legend transparency .xlsx | column chart legend clear background Aspose.Cells | Aspose.Cells legend entry IsTextNoFill

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, inserts a column chart, sets the chart's Legend.BackgroundMode to Transparent, iterates through each LegendEntry to also set BackgroundMode to Transparent and IsTextNoFill to true, and saves the file as TransparentLegend.xlsx.
class Program
{
    static void Main()
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

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Set the legend background to transparent
        chart.Legend.BackgroundMode = BackgroundMode.Transparent;

        // Ensure each legend entry has no fill (transparent background and no text fill)
        foreach (LegendEntry entry in chart.Legend.LegendEntries)
        {
            entry.BackgroundMode = BackgroundMode.Transparent; // entry background transparent
            entry.IsTextNoFill = true;                         // text has no fill
        }

        // Save the workbook
        workbook.Save("TransparentLegend.xlsx");
    }
}
