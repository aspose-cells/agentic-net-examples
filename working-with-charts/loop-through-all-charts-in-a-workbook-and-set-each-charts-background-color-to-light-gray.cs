using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class SetChartBackgroundColor
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // -------------------------------------------------
        // Sample data and chart creation (optional demo)
        // -------------------------------------------------
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["A4"].PutValue("Banana");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(50);
        sheet.Cells["B3"].PutValue(30);
        sheet.Cells["B4"].PutValue(20);

        // Add a chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // -------------------------------------------------
        // Loop through all worksheets and their charts
        // Set each chart's background color to LightGray
        // -------------------------------------------------
        foreach (Worksheet ws in workbook.Worksheets)
        {
            foreach (Chart ch in ws.Charts)
            {
                // Set the background color of the chart area
                ch.ChartArea.Area.BackgroundColor = Color.LightGray;

                // Optionally, also set the plot area background to keep consistency
                ch.PlotArea.Area.BackgroundColor = Color.LightGray;
            }
        }

        // Save the workbook
        workbook.Save("ChartsWithLightGrayBackground.xlsx");
    }
}