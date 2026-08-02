using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class CombinedColumnLineChart
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        // Categories
        sheet.Cells["A1"].PutValue("Month");
        sheet.Cells["A2"].PutValue("Jan");
        sheet.Cells["A3"].PutValue("Feb");
        sheet.Cells["A4"].PutValue("Mar");
        sheet.Cells["A5"].PutValue("Apr");

        // Column series data (e.g., Sales)
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["B4"].PutValue(180);
        sheet.Cells["B5"].PutValue(200);

        // Line series data (e.g., Profit)
        sheet.Cells["C1"].PutValue("Profit");
        sheet.Cells["C2"].PutValue(30);
        sheet.Cells["C3"].PutValue(45);
        sheet.Cells["C4"].PutValue(50);
        sheet.Cells["C5"].PutValue(70);

        // Add a chart of type Column (primary axis will host column series)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Add the column series (primary axis)
        chart.NSeries.Add("B2:B5", true);          // Values
        chart.NSeries[0].Name = "Sales";          // Series name
        chart.NSeries.CategoryData = "A2:A5";     // Categories

        // Add the line series (secondary axis)
        chart.NSeries.Add("C2:C5", true);          // Values
        chart.NSeries[1].Name = "Profit";         // Series name

        // Change the second series to a line type
        chart.NSeries[1].Type = ChartType.Line;

        // Plot the second series on the secondary value axis
        chart.NSeries[1].PlotOnSecondAxis = true;

        // Optional: customize axes titles
        chart.CategoryAxis.Title.Text = "Month";
        chart.ValueAxis.Title.Text = "Sales";
        chart.SecondValueAxis.Title.Text = "Profit";

        // Save the workbook
        workbook.Save("CombinedColumnLineChart.xlsx");
    }
}