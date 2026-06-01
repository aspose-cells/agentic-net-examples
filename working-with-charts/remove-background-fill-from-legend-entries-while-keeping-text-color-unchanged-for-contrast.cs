using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class RemoveLegendBackground
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
        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["B2"].PutValue(50);
        sheet.Cells["B3"].PutValue(80);
        sheet.Cells["C1"].PutValue("Series2");
        sheet.Cells["C2"].PutValue(70);
        sheet.Cells["C3"].PutValue(60);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:C3", true);          // Add series data
        chart.NSeries.CategoryData = "A2:A3";      // Set category (X‑axis) data
        chart.ShowLegend = true;                   // Ensure the legend is displayed

        // Remove background fill from each legend entry while keeping text color unchanged
        foreach (LegendEntry entry in chart.Legend.LegendEntries)
        {
            entry.IsTextNoFill = true; // Disables background fill of the legend text
            // No change to entry.Font.Color, so text color stays the same for contrast
        }

        // Save the workbook
        workbook.Save("ChartWithoutLegendBackground.xlsx");
    }
}