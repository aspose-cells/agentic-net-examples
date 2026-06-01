using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

class RichTextDataLabelDemo
{
    public static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a column chart
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels for the first series
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = true;

        // Access the first data point and its data label
        ChartPoint point = series.Points[0];
        DataLabels dataLabel = point.DataLabels;

        // Set custom text for the data label
        dataLabel.Text = "10 units";

        // Apply mixed font sizes:
        // "10" (characters 0-2) larger and blue
        dataLabel.Characters(0, 2).Font.Size = 16;
        dataLabel.Characters(0, 2).Font.Color = Color.Blue;

        // " units" (remaining characters) smaller and gray
        int remainingLength = dataLabel.Text.Length - 2;
        dataLabel.Characters(2, remainingLength).Font.Size = 10;
        dataLabel.Characters(2, remainingLength).Font.Color = Color.Gray;

        // Apply the font settings to all child nodes (optional but ensures consistency)
        dataLabel.ApplyFont();

        // Save the workbook
        workbook.Save("RichTextDataLabel.xlsx");
    }
}