using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class EnableAutomaticUnits
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data with large numbers
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(1500000);
        sheet.Cells["B3"].PutValue(3000000);
        sheet.Cells["B4"].PutValue(4500000);

        // Insert a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the chart data range
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable automatic display units on the Y‑axis
        // Setting DisplayUnit to None lets Excel choose the appropriate unit.
        // Show the unit label if a unit is applied.
        chart.ValueAxis.DisplayUnit = DisplayUnitType.None;
        chart.ValueAxis.IsDisplayUnitLabelShown = true;

        // Save the workbook
        workbook.Save("AutomaticUnitsChart.xlsx", SaveFormat.Xlsx);
    }
}