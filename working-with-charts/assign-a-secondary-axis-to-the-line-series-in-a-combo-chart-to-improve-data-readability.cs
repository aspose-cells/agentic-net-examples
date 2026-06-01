using Aspose.Cells;
using Aspose.Cells.Charts;

class ComboChartSecondaryAxisDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        sheet.Cells["A1"].PutValue("Month");
        sheet.Cells["A2"].PutValue("Jan");
        sheet.Cells["A3"].PutValue("Feb");
        sheet.Cells["A4"].PutValue("Mar");
        sheet.Cells["A5"].PutValue("Apr");

        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["B2"].PutValue(200);
        sheet.Cells["B3"].PutValue(250);
        sheet.Cells["B4"].PutValue(300);
        sheet.Cells["B5"].PutValue(280);

        sheet.Cells["C1"].PutValue("Profit");
        sheet.Cells["C2"].PutValue(20);
        sheet.Cells["C3"].PutValue(30);
        sheet.Cells["C4"].PutValue(25);
        sheet.Cells["C5"].PutValue(35);

        // Add a combo chart (Column + Line)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
        Chart chart = sheet.Charts[chartIndex];

        // First series (column) – Sales
        chart.NSeries.Add("B2:B5", true);
        // Second series (line) – Profit
        chart.NSeries.Add("C2:C5", true);
        // Set category (X) axis data
        chart.NSeries.CategoryData = "A2:A5";

        // Change the second series type to Line
        chart.NSeries[1].Type = ChartType.Line;

        // Plot the line series on the secondary Y axis
        chart.NSeries[1].PlotOnSecondAxis = true;

        // Optional: customize the secondary axis title
        Axis secondaryAxis = chart.SecondValueAxis;
        secondaryAxis.Title.Text = "Profit (Secondary Axis)";

        // Save the workbook
        workbook.Save("ComboChartSecondaryAxis.xlsx");
    }
}