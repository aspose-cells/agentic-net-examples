using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SetDataLabelAlignment
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for a horizontal bar chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");

        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a horizontal bar chart (ChartType.Bar)
        int chartIndex = sheet.Charts.Add(ChartType.Bar, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Enable data labels for the first series
        chart.NSeries[0].DataLabels.ShowValue = true;

        // Right‑justify the data label text
        chart.NSeries[0].DataLabels.TextHorizontalAlignment = TextAlignmentType.Right;

        // Save the workbook
        workbook.Save("HorizontalBarChart_WithRightAlignedDataLabels.xlsx");
    }
}