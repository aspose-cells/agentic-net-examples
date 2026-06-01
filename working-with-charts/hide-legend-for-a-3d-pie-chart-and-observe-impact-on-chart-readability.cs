using Aspose.Cells;
using Aspose.Cells.Charts;

class HideLegend3DPieChart
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pie chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(30);
        worksheet.Cells["B3"].PutValue(50);
        worksheet.Cells["B4"].PutValue(20);

        // Add a 3‑D pie chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Pie3D, 5, 0, 20, 15);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Hide the legend to see how it affects readability
        chart.ShowLegend = false;

        // Add a title indicating that the legend is hidden
        chart.Title.Text = "3‑D Pie Chart (Legend Hidden)";

        // Save the workbook
        workbook.Save("HideLegend3DPieChart.xlsx");
    }
}