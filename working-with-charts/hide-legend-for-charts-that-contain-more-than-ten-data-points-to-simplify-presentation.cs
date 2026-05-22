using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class HideLegendForLargeCharts
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data with more than ten data points
        sheet.Cells["A1"].PutValue("Item");
        sheet.Cells["B1"].PutValue("Value");
        for (int i = 2; i <= 13; i++)
        {
            sheet.Cells[$"A{i}"].PutValue($"Item{i - 1}");
            sheet.Cells[$"B{i}"].PutValue((i - 1) * 10);
        }

        // Add a column chart that uses the data range
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B13", true);          // values
        chart.NSeries.CategoryData = "A2:A13";     // categories

        // Determine the number of data points in the first series.
        // The Values property returns the range string (e.g., "B2:B13").
        // We parse the range to calculate the count.
        string range = chart.NSeries[0].Values;    // e.g., "B2:B13"
        int dataPointCount = 0;
        if (!string.IsNullOrEmpty(range))
        {
            // Simple parsing assuming a single contiguous range like "B2:B13"
            string[] parts = range.Split(':');
            if (parts.Length == 2)
            {
                // Extract row numbers
                int startRow = int.Parse(System.Text.RegularExpressions.Regex.Match(parts[0], @"\d+").Value);
                int endRow   = int.Parse(System.Text.RegularExpressions.Regex.Match(parts[1], @"\d+").Value);
                dataPointCount = Math.Abs(endRow - startRow) + 1;
            }
        }

        // Hide the legend if the chart contains more than ten data points
        if (dataPointCount > 10)
        {
            chart.ShowLegend = false;   // using Chart.ShowLegend property
        }

        // Save the workbook
        workbook.Save("ChartHideLegend.xlsx");
    }
}