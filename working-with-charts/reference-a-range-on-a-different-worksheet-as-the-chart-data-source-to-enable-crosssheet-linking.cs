using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class CrossSheetChartDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Add a worksheet that will hold the source data
        Worksheet dataSheet = wb.Worksheets.Add("DataSheet");

        // Populate sample data in the data sheet
        dataSheet.Cells["A1"].PutValue("Category");
        dataSheet.Cells["B1"].PutValue("Value");
        dataSheet.Cells["A2"].PutValue("A");
        dataSheet.Cells["B2"].PutValue(10);
        dataSheet.Cells["A3"].PutValue("B");
        dataSheet.Cells["B3"].PutValue(20);
        dataSheet.Cells["A4"].PutValue("C");
        dataSheet.Cells["B4"].PutValue(30);

        // Access the first worksheet where the chart will be placed
        Worksheet chartSheet = wb.Worksheets[0];
        chartSheet.Name = "ChartSheet";

        // Add a column chart to the chart sheet
        int chartIndex = chartSheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = chartSheet.Charts[chartIndex];

        // Set the chart's data range to a range on a different worksheet (cross‑sheet reference)
        // Note the sheet name must be enclosed in single quotes if it contains spaces or special characters
        chart.SetChartDataRange("'DataSheet'!$A$1:$B$4", true);

        // Optionally link the chart title to a cell on the data sheet
        dataSheet.Cells["C1"].PutValue("Sales Chart");
        chart.Title.LinkedSource = "'DataSheet'!$C$1";

        // Save the workbook
        wb.Save("CrossSheetChart.xlsx");
    }
}