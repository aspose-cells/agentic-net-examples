using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class SetDataLabelShapeEllipse
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the bubble chart
        worksheet.Cells["A1"].PutValue("X");
        worksheet.Cells["A2"].PutValue(1);
        worksheet.Cells["A3"].PutValue(2);
        worksheet.Cells["A4"].PutValue(3);

        worksheet.Cells["B1"].PutValue("Y");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        worksheet.Cells["C1"].PutValue("Size");
        worksheet.Cells["C2"].PutValue(5);
        worksheet.Cells["C3"].PutValue(10);
        worksheet.Cells["C4"].PutValue(15);

        // Add a bubble chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Bubble, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Add a series to the chart and set X, Y and bubble size ranges
        int seriesIndex = chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";
        Series series = chart.NSeries[seriesIndex];
        series.BubbleSizes = "C2:C4";

        // Enable data labels and set their shape type to ellipse
        series.DataLabels.ShowValue = true;
        series.DataLabels.ShapeType = DataLabelShapeType.Ellipse;

        // Save the workbook with the configured chart
        workbook.Save("BubbleChartDataLabelEllipse.xlsx");
    }
}