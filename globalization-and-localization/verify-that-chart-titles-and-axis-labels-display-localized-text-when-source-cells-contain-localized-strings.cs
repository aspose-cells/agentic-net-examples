using Aspose.Cells;
using Aspose.Cells.Charts;

class VerifyChartLocalization
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Put localized strings into cells (e.g., Chinese)
        worksheet.Cells["A1"].PutValue("图表标题");      // Chart title
        worksheet.Cells["B1"].PutValue("类别轴标题");   // Category (X) axis title
        worksheet.Cells["C1"].PutValue("数值轴标题");   // Value (Y) axis title

        // Add sample data for the chart
        worksheet.Cells["A2"].PutValue("一");
        worksheet.Cells["A3"].PutValue("二");
        worksheet.Cells["A4"].PutValue("三");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = worksheet.Charts[chartIndex];

        // Bind data series and categories
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Set chart title and axis titles using the localized cell values
        chart.Title.Text = worksheet.Cells["A1"].StringValue;
        chart.CategoryAxis.Title.Text = worksheet.Cells["B1"].StringValue;
        chart.ValueAxis.Title.Text = worksheet.Cells["C1"].StringValue;

        // Save the workbook (lifecycle rule)
        workbook.Save("LocalizedChart.xlsx");
    }
}