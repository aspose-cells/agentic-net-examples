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
        // Column A – Category names, Column B – Values (positive for increase, negative for decrease)
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");

        string[] categories = { "Start", "Revenue", "Expense", "Profit" };
        double[] values = { 0, 500, -200, 300 };

        for (int i = 0; i < categories.Length; i++)
        {
            sheet.Cells[i + 1, 0].PutValue(categories[i]);   // A2, A3, ...
            sheet.Cells[i + 1, 1].PutValue(values[i]);      // B2, B3, ...
        }

        // Add a Waterfall chart to the worksheet (using the provided Add method)
        int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Bind the data range to the chart
        chart.NSeries.Add("B2:B5", true);          // Values
        chart.NSeries.CategoryData = "A2:A5";      // Categories

        // Designate the last data point (index 3) as a total segment
        // The Subtotals property receives an array of zero‑based indices
        chart.NSeries[0].LayoutProperties.Subtotals = new int[] { 3 };

        // Optional: set a chart title
        chart.Title.Text = "Waterfall Example";

        // Save the workbook with the chart
        workbook.Save("WaterfallChart.xlsx", SaveFormat.Xlsx);
    }
}