using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class DuplicateChartExample
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

        // Define the data range string (used for both charts)
        string dataRange = "A1:B4";

        // -------------------------------------------------
        // Add the original chart
        // -------------------------------------------------
        int originalChartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart originalChart = sheet.Charts[originalChartIndex];
        originalChart.SetChartDataRange(dataRange, true);
        // Set legend position for the original chart (default is Right)
        originalChart.Legend.Position = LegendPositionType.Right;

        // -------------------------------------------------
        // Duplicate the chart
        // -------------------------------------------------
        // Add a new chart with the same type; place it below the original
        int duplicateChartIndex = sheet.Charts.Add(ChartType.Column, 16, 0, 26, 5);
        Chart duplicateChart = sheet.Charts[duplicateChartIndex];
        duplicateChart.SetChartDataRange(dataRange, true);
        // Change the legend position for the duplicated chart
        duplicateChart.Legend.Position = LegendPositionType.Left;

        // Optionally demonstrate moving the duplicated chart to a specific range
        duplicateChart.Move(16, 0, 26, 5);

        // Save the workbook containing both charts
        workbook.Save("DuplicatedChart.xlsx");
    }
}