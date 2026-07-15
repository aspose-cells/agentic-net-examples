using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class MovingAverageChartDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // Populate sample data in column A (1‑12)
            for (int i = 0; i < 12; i++)
            {
                cells[i, 0].PutValue(i + 1); // A1..A12 = 1..12
            }

            int period = 3; // Moving‑average period

            // Fill column B with moving‑average formulas
            for (int row = 0; row < 12; row++)
            {
                if (row + 1 < period)
                {
                    // For the first rows where a full window is not available, copy the original value
                    cells[row, 1].PutValue(cells[row, 0].Value);
                }
                else
                {
                    // Build formula: =AVERAGE(A{start}:A{end})
                    int startRow = row - period + 2; // Excel rows are 1‑based
                    int endRow = row + 1;
                    string formula = $"=AVERAGE(A{startRow}:A{endRow})";

                    // Set the formula using the Formula property (compatible with all versions)
                    cells[row, 1].Formula = formula;
                }
            }

            // Add a line chart to display the moving‑average series
            int chartIndex = ws.Charts.Add(ChartType.Line, 5, 0, 20, 12);
            Chart chart = ws.Charts[chartIndex];

            // Use column B (moving averages) as the series values and column A as categories
            chart.NSeries.Add("B1:B12", true);
            chart.NSeries.CategoryData = "A1:A12";

            // Optionally add a moving‑average trendline to the series (same period)
            int trendlineIdx = chart.NSeries[0].TrendLines.Add(TrendlineType.MovingAverage);
            chart.NSeries[0].TrendLines[trendlineIdx].Period = period;

            // Calculate all formulas so the chart reflects the latest data
            wb.CalculateFormula();

            // Save the workbook
            wb.Save("MovingAverageChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}