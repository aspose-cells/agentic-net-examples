using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // (Optional) Add some sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a chart (initial position does not matter, it will be moved later)
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 0, 0, 10, 5);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data source for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Resize the chart: width = 500 points, height = 300 points
        chart.ChartObject.Width = 500;
        chart.ChartObject.Height = 300;

        // Position the chart so its upper‑left corner aligns with cell D5
        // Row and column indices are zero‑based (D = 3, 5th row = 4)
        chart.Move(4, 3, 4, 3);

        // Save the workbook with the resized and repositioned chart
        workbook.Save("ResizedChart.xlsx");
    }
}