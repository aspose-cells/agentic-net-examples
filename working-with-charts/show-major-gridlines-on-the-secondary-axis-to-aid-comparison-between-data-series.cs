using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

class ShowSecondaryAxisMajorGridlines
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");

        worksheet.Cells["B1"].PutValue("Series 1");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        worksheet.Cells["C1"].PutValue("Series 2");
        worksheet.Cells["C2"].PutValue(200);
        worksheet.Cells["C3"].PutValue(400);
        worksheet.Cells["C4"].PutValue(600);

        // Add a column chart
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = worksheet.Charts[chartIndex];

        // Add two series: first on primary axis, second on secondary axis
        chart.NSeries.Add("B2:B4", true); // primary series
        chart.NSeries.Add("C2:C4", true); // secondary series
        chart.NSeries.CategoryData = "A2:A4";

        // Plot the second series on the secondary value axis
        chart.NSeries[1].PlotOnSecondAxis = true;

        // Enable major gridlines on the secondary value axis
        chart.SecondValueAxis.MajorGridLines.IsVisible = true;
        // Optional: set gridline color for better visibility
        chart.SecondValueAxis.MajorGridLines.Color = Color.Blue;

        // Save the workbook
        workbook.Save("SecondaryAxisMajorGridlines.xlsx");
    }
}