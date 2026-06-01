using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class ComboChartExample
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        // Categories
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Jan");
        sheet.Cells["A3"].PutValue("Feb");
        sheet.Cells["A4"].PutValue("Mar");

        // Column series values
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["B4"].PutValue(180);

        // Line series values
        sheet.Cells["C1"].PutValue("Profit");
        sheet.Cells["C2"].PutValue(30);
        sheet.Cells["C3"].PutValue(45);
        sheet.Cells["C4"].PutValue(60);

        // Area series values
        sheet.Cells["D1"].PutValue("Expenses");
        sheet.Cells["D2"].PutValue(70);
        sheet.Cells["D3"].PutValue(80);
        sheet.Cells["D4"].PutValue(90);

        // Add a combo chart (primary type can be Column; using 3‑D to expose Series axis)
        int chartIndex = sheet.Charts.Add(ChartType.Column3DClustered, 6, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set chart title
        chart.Title.Text = "Combo Chart: Column, Line, Area";

        // Add the three series (rule: SeriesCollection.Add)
        // Series 0 – Column (primary value axis)
        chart.NSeries.Add("=Sheet1!$B$2:$B$4", true);
        chart.NSeries[0].Name = "Sales";
        chart.NSeries[0].Type = ChartType.Column;

        // Series 1 – Line (secondary value axis)
        chart.NSeries.Add("=Sheet1!$C$2:$C$4", true);
        chart.NSeries[1].Name = "Profit";
        chart.NSeries[1].Type = ChartType.Line;
        chart.NSeries[1].PlotOnSecondAxis = true; // use secondary axis

        // Series 2 – Area (series axis – available on 3‑D charts)
        chart.NSeries.Add("=Sheet1!$D$2:$D$4", true);
        chart.NSeries[2].Name = "Expenses";
        chart.NSeries[2].Type = ChartType.Area;
        // For 3‑D charts, the Series axis is distinct; we associate the series with it
        // (Aspose.Cells automatically maps Area series to the series axis in 3‑D combo charts)

        // Set category (X) data for all series
        chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$4";

        // Customize axes titles to illustrate distinct axes
        chart.ValueAxis.Title.Text = "Primary Value Axis (Sales)";
        chart.SecondValueAxis.Title.Text = "Secondary Value Axis (Profit)";
        chart.SeriesAxis.Title.Text = "Series Axis (Expenses)";

        // Save the workbook (lifecycle rule: save)
        workbook.Save("ComboChart.xlsx", SaveFormat.Xlsx);
    }
}