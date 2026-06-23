using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Enable3DChartDepth
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["A3"].PutValue("Q2");
        sheet.Cells["A4"].PutValue("Q3");

        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["B4"].PutValue(180);

        // Add a 3‑D column chart (this enables 3‑D formatting)
        int chartIndex = sheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Set the depth of the 3‑D chart.
        // DepthPercent defines the depth as a percentage of the chart width.
        // Here we set it to 150% (approximately fifteen points relative to the default width).
        chart.DepthPercent = 150;

        // Save the workbook
        workbook.Save("ChartWith3DDepth.xlsx");
    }
}