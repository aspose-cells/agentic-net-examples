using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class TransparentLegendDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["A3"].PutValue("Q2");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(50);
        sheet.Cells["B3"].PutValue(80);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B3", true);          // Data values
        chart.NSeries.CategoryData = "A2:A3";      // Category labels

        // Set the legend background to transparent
        chart.Legend.BackgroundMode = BackgroundMode.Transparent;

        // Ensure each legend entry has no fill (text fill disabled)
        foreach (LegendEntry entry in chart.Legend.LegendEntries)
        {
            entry.IsTextNoFill = true;
        }

        // Save the workbook
        workbook.Save("TransparentLegend.xlsx");
    }
}