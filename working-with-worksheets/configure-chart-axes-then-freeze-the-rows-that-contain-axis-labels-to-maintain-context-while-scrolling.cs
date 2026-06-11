using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data: categories in column A, values in column B
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue("Q1");
        worksheet.Cells["A3"].PutValue("Q2");
        worksheet.Cells["A4"].PutValue("Q3");
        worksheet.Cells["B2"].PutValue(120);
        worksheet.Cells["B3"].PutValue(150);
        worksheet.Cells["B4"].PutValue(180);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Configure the Category (X) axis
        chart.CategoryAxis.Title.Text = "Quarter";
        chart.CategoryAxis.Title.IsVisible = true;
        chart.CategoryAxis.MajorTickMark = TickMarkType.Outside;
        chart.CategoryAxis.MinorTickMark = TickMarkType.Inside;

        // Configure the Value (Y) axis
        chart.ValueAxis.Title.Text = "Sales";
        chart.ValueAxis.Title.IsVisible = true;
        chart.ValueAxis.IsAutomaticMinValue = false;
        chart.ValueAxis.MinValue = 0;
        chart.ValueAxis.IsAutomaticMaxValue = false;
        chart.ValueAxis.MaxValue = 200;
        chart.ValueAxis.MajorUnit = 50;
        chart.ValueAxis.MajorTickMark = TickMarkType.Outside;
        chart.ValueAxis.MinorTickMark = TickMarkType.Inside;

        // Freeze the row that contains the axis labels (row 1)
        // Freeze at cell A2 with 1 frozen row and 0 frozen columns
        worksheet.FreezePanes("A2", 1, 0);

        // Save the workbook
        workbook.Save("ChartWithFrozenAxisLabels.xlsx");
    }
}