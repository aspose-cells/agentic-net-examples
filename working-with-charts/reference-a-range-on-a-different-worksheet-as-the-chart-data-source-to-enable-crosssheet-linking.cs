using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class CrossSheetChartDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // ---------- Source worksheet with data ----------
        Worksheet sourceSheet = workbook.Worksheets[0];
        sourceSheet.Name = "SourceData";

        // Populate sample data
        sourceSheet.Cells["A1"].PutValue("Category");
        sourceSheet.Cells["B1"].PutValue("Value");
        sourceSheet.Cells["A2"].PutValue("A");
        sourceSheet.Cells["B2"].PutValue(10);
        sourceSheet.Cells["A3"].PutValue("B");
        sourceSheet.Cells["B3"].PutValue(20);
        sourceSheet.Cells["A4"].PutValue("C");
        sourceSheet.Cells["B4"].PutValue(30);

        // Optional: place a title in the source sheet to link the chart title
        sourceSheet.Cells["D1"].PutValue("Sales Chart");

        // ---------- Destination worksheet for the chart ----------
        Worksheet chartSheet = workbook.Worksheets.Add("ChartSheet");

        // Add a column chart to the destination sheet
        int chartIndex = chartSheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = chartSheet.Charts[chartIndex];

        // Set the chart data range to reference the source worksheet
        // The range string includes the sheet name, e.g., "SourceData!$A$1:$B$4"
        chart.SetChartDataRange("SourceData!$A$1:$B$4", true);

        // Link the chart title to the cell on the source sheet (optional)
        chart.Title.LinkedSource = "='SourceData'!$D$1";

        // Save the workbook
        workbook.Save("CrossSheetChart.xlsx");
    }
}