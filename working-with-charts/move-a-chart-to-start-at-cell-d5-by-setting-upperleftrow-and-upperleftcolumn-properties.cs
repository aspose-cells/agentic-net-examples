using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System;

class MoveChartExample
{
    static void Main()
    {
        // Create a new workbook
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

        // Add a chart (initial position does not matter)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 0, 0, 10, 5);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Move the chart so its upper‑left corner starts at cell D5
        // D5 corresponds to row index 4 and column index 3 (zero‑based)
        chart.ChartObject.UpperLeftRow = 4;      // Row 5
        chart.ChartObject.UpperLeftColumn = 3;   // Column D

        // Save the workbook
        workbook.Save("ChartMoved.xlsx");
    }
}