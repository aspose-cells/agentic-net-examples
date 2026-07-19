// Title: Hide Aspose.Cells chart legend when data points exceed 10 (C#)
// Description: Creates a workbook, fills columns A‑B with 12 items, adds a column chart, counts the total data points across all series, and disables the legend if the count is greater than ten before saving the file.
// Keywords: Aspose.Cells hide legend | C# chart legend conditional | count chart data points Aspose.Cells | ShowLegend property | Aspose.Cells chart customization
// Common Searches: Aspose.Cells hide legend for large chart | C# count data points in Aspose.Cells chart | conditional legend visibility Aspose.Cells .NET | how to disable chart legend based on series size | Aspose.Cells chart ShowLegend example
// Developer Intent: Automatically suppress the chart legend when the chart contains more than ten data points.
// Use Cases: Generate column charts for reports and hide legends on charts with many categories to keep the layout tidy. | Apply a uniform rule across multiple financial dashboards that removes legends from dense charts for better print readability. | Create exportable workbooks where charts with over ten data points automatically hide their legends to improve visual clarity.
// AI Prompts: Provide C# code using Aspose.Cells that counts all data points in a chart and sets ShowLegend = false when the total exceeds a threshold. | Show how to iterate over chart series, derive row and column counts from the series values range, and toggle legend visibility. | Explain the steps to conditionally hide a chart legend in Aspose.Cells based on the number of categories or series points.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeRange = Aspose.Cells.Range;

// Creates a workbook, fills columns A‑B with 12 items, adds a column chart, counts the total data points across all series, and disables the legend if the count is greater than ten before saving the file.
class HideLegendForLargeCharts
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with more than ten data points
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            for (int i = 2; i <= 13; i++) // 12 data points (rows 2‑13)
            {
                sheet.Cells[$"A{i}"].PutValue($"Item{i - 1}");
                sheet.Cells[$"B{i}"].PutValue(i * 10);
            }

            // Add a column chart that uses the data range
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B13", true);          // Values
            chart.NSeries.CategoryData = "A2:A13";     // Categories

            // Determine the total number of data points in the chart
            int totalPoints = 0;
            foreach (Series series in chart.NSeries)
            {
                if (!string.IsNullOrEmpty(series.Values))
                {
                    // Create a range from the series values string (handles full address formats)
                    AsposeRange dataRange = sheet.Cells.CreateRange(series.Values);
                    totalPoints += dataRange.RowCount * dataRange.ColumnCount;
                }
            }

            // Hide the legend if the chart contains more than ten data points
            if (totalPoints > 10)
            {
                chart.ShowLegend = false;
            }

            // Save the workbook
            workbook.Save("ChartWithConditionalLegend.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
