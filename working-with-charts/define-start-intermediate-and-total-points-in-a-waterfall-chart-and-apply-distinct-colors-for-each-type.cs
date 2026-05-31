using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

class WaterfallChartDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Prepare sample data for the waterfall chart
        // Column A: Category names, Column B: Values
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");

        string[] categories = { "Start", "Q1", "Q2", "Q3", "Total" };
        double[] values = { 0, 30, -10, 20, 0 }; // placeholder values; actual totals are calculated by the chart

        for (int i = 0; i < categories.Length; i++)
        {
            sheet.Cells[i + 2, 0].PutValue(categories[i]); // A column
            sheet.Cells[i + 2, 1].PutValue(values[i]);    // B column
        }

        // Add a Waterfall chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the series and the category axis
        chart.NSeries.Add("B3:B7", true);          // Values range (B3:B7)
        chart.NSeries.CategoryData = "A3:A7";     // Categories range (A3:A7)

        // Define distinct colors for start, intermediate, and total points
        Color startColor = Color.LightBlue;
        Color intermediateColor = Color.LightGreen;
        Color totalColor = Color.Orange;

        // Apply colors to the points
        // First point is the start point
        chart.NSeries[0].Points[0].Area.ForegroundColor = startColor;

        // Intermediate points (excluding the last point)
        for (int i = 1; i < chart.NSeries[0].Points.Count - 1; i++)
        {
            chart.NSeries[0].Points[i].Area.ForegroundColor = intermediateColor;
        }

        // Last point is the total point
        int lastIndex = chart.NSeries[0].Points.Count - 1;
        chart.NSeries[0].Points[lastIndex].Area.ForegroundColor = totalColor;

        // Save the workbook with the chart
        workbook.Save("WaterfallChartDemo.xlsx");
    }
}