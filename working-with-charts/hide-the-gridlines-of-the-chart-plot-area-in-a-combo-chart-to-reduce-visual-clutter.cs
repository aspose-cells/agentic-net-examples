using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class HideChartGridlines
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the combo chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["C1"].PutValue("Series2");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);
        sheet.Cells["C2"].PutValue(15);
        sheet.Cells["C3"].PutValue(25);
        sheet.Cells["C4"].PutValue(35);

        // Add a combo chart (column + line) to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // First series as column
        chart.NSeries.Add("B2:B4", true);
        // Second series as line (combo)
        chart.NSeries.Add("C2:C4", true);
        chart.NSeries[1].Type = ChartType.Line;

        // Set category (X‑axis) data
        chart.NSeries.CategoryData = "A2:A4";

        // Hide all gridlines in the plot area to reduce visual clutter
        chart.ValueAxis.MajorGridLines.IsVisible = false;
        chart.ValueAxis.MinorGridLines.IsVisible = false;
        chart.CategoryAxis.MajorGridLines.IsVisible = false;
        chart.CategoryAxis.MinorGridLines.IsVisible = false;

        // Save the workbook with the chart
        workbook.Save("ComboChart_NoGridlines.xlsx", SaveFormat.Xlsx);
    }
}