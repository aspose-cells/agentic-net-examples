using Aspose.Cells;
using Aspose.Cells.Charts;

class SetSecondaryAxisPercentage
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");

        sheet.Cells["B1"].PutValue("Primary");
        sheet.Cells["B2"].PutValue(0.2);
        sheet.Cells["B3"].PutValue(0.4);
        sheet.Cells["B4"].PutValue(0.6);

        sheet.Cells["C1"].PutValue("Secondary");
        sheet.Cells["C2"].PutValue(0.1);
        sheet.Cells["C3"].PutValue(0.3);
        sheet.Cells["C4"].PutValue(0.5);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Add two series: first on primary axis, second on secondary axis
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.Add("C2:C4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Plot the second series on the secondary value axis
        chart.NSeries[1].PlotOnSecondAxis = true;

        // Set custom number format for secondary axis tick labels to display percentages
        chart.SecondValueAxis.TickLabels.NumberFormat = "0%";

        // Save the workbook
        workbook.Save("SecondaryAxisPercentage.xlsx");
    }
}