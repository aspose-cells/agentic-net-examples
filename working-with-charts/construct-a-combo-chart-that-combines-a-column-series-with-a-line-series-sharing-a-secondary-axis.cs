using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class ComboChartExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data
        worksheet.Cells["A1"].PutValue("Month");
        worksheet.Cells["A2"].PutValue("Jan");
        worksheet.Cells["A3"].PutValue("Feb");
        worksheet.Cells["A4"].PutValue("Mar");

        worksheet.Cells["B1"].PutValue("Sales");
        worksheet.Cells["B2"].PutValue(120);
        worksheet.Cells["B3"].PutValue(150);
        worksheet.Cells["B4"].PutValue(180);

        worksheet.Cells["C1"].PutValue("Profit");
        worksheet.Cells["C2"].PutValue(30);
        worksheet.Cells["C3"].PutValue(45);
        worksheet.Cells["C4"].PutValue(60);

        // Add a column chart (will serve as the primary series)
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // First series: Sales (column)
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Second series: Profit (line)
        chart.NSeries.Add("C2:C4", true);
        // Change the series type to Line
        chart.NSeries[1].Type = ChartType.Line;
        // Plot this series on the secondary Y axis
        chart.NSeries[1].PlotOnSecondAxis = true;

        // Optional: give the secondary axis a title
        chart.SecondValueAxis.Title.Text = "Profit";

        // Save the workbook with the combo chart
        workbook.Save("ComboChart.xlsx");
    }
}