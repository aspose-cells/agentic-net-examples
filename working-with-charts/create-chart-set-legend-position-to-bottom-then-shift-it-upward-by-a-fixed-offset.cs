using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
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

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Set the legend position to the bottom of the chart
        chart.Legend.Position = LegendPositionType.Bottom;

        // Shift the legend upward by a fixed offset (e.g., 5% of the chart height)
        // YRatioToChart is a fraction (0‑1) representing the vertical position relative to the chart area.
        // Decreasing the value moves the legend upward.
        chart.Legend.YRatioToChart -= 0.05; // move up by 5%

        // Save the workbook with the modified chart
        workbook.Save("ChartWithShiftedLegend.xlsx");
    }
}