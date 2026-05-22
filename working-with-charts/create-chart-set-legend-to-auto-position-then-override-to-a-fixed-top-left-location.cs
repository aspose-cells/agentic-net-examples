using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class ChartLegendPositionDemo
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

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Access the legend
        Legend legend = chart.Legend;

        // 1. Set legend to automatic positioning
        legend.SetPositionAuto(); // legend now uses default docking (right side)

        // 2. Override to a fixed top‑left location
        //    Use NotDocked to allow manual placement, then set ratios relative to the chart area
        legend.Position = LegendPositionType.NotDocked;
        legend.XRatioToChart = 0.02; // 2% from the left edge of the chart
        legend.YRatioToChart = 0.02; // 2% from the top edge of the chart
        legend.WidthRatioToChart = 0.2; // 20% of chart width
        legend.HeightRatioToChart = 0.1; // 10% of chart height

        // Optional: customize appearance
        legend.Font.Size = 12;
        legend.Font.IsBold = true;
        legend.IsOverLay = false; // legend will not overlap the plot area

        // Save the workbook
        workbook.Save("ChartLegendPositionDemo.xlsx");
    }
}