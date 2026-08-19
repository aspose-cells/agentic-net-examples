// Title: Aspose.Cells for .NET: Line chart with moving‑average series using a formula column
// Description: Creates a workbook, fills columns A (categories) and B (values), inserts an AVERAGE formula in column C to compute a 3‑point moving average, forces formula calculation, adds a line chart, binds original and moving‑average series, sets the X‑axis to column A, applies basic formatting, and saves as MovingAverageChart.xlsx.
// Keywords: Aspose.Cells | C# | .NET | line chart | moving average | formula column | Excel AVERAGE function | chart series from formula | dynamic chart data | calculate rolling average
// Common Searches: Aspose.Cells line chart with moving average | use formula column as chart series Aspose.Cells .NET | calculate rolling average in Excel with Aspose.Cells | add moving average series to Excel chart C# | chart series from calculated cells Aspose.Cells
// Developer Intent: Generate a line chart that displays both raw values and a calculated moving‑average series derived from an Excel formula column.
// Use Cases: Sales or revenue trend reports where a moving average smooths monthly volatility. | Financial dashboards that overlay rolling averages on key performance indicators. | Production or inventory charts that automatically update the average line as new data is added.
// AI Prompts: Show how to make the moving‑average period configurable at runtime and refresh the chart. | Provide code to style the moving‑average series (e.g., dashed line, custom color) in the generated chart. | Explain how to replace hard‑coded ranges with named ranges for both original and moving‑average series.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsMovingAverageChart
{
    // Creates a workbook, fills columns A (categories) and B (values), inserts an AVERAGE formula in column C to compute a 3‑point moving average, forces formula calculation, adds a line chart, binds original and moving‑average series, sets the X‑axis to column A, applies basic formatting, and saves as MovingAverageChart.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            // Column A: Category (e.g., Month)
            // Column B: Original values
            string[] categories = { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct" };
            double[] values = { 120, 150, 130, 170, 160, 180, 200, 190, 210, 230 };
            int rowCount = categories.Length;

            for (int i = 0; i < rowCount; i++)
            {
                cells[i + 1, 0].PutValue(categories[i]); // A column (index 0)
                cells[i + 1, 1].PutValue(values[i]);    // B column (index 1)
            }

            // Define moving average period
            int period = 3;

            // Column C will hold the moving average calculated by a formula
            // For rows where there are not enough previous points, leave the cell empty
            for (int i = 0; i < rowCount; i++)
            {
                int currentRow = i + 1; // Excel rows are 1‑based
                if (i + 1 >= period)
                {
                    // Formula: =AVERAGE(B{row-period+1}:B{row})
                    string formula = $"=AVERAGE(B{currentRow - period + 1}:B{currentRow})";
                    cells[currentRow, 2].Formula = formula; // C column (index 2)
                }
                else
                {
                    cells[currentRow, 2].PutValue(string.Empty);
                }
            }

            // Calculate all formulas so that the moving average values are materialized
            workbook.CalculateFormula();

            // Add a line chart
            int chartIndex = sheet.Charts.Add(ChartType.Line, 12, 0, 30, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Series 1: Original values (B column)
            chart.NSeries.Add("B2:B11", true);
            chart.NSeries[0].Name = "Original";

            // Series 2: Moving average (C column)
            chart.NSeries.Add("C2:C11", true);
            chart.NSeries[1].Name = "Moving Average";

            // Set category (X) axis data (A column)
            chart.NSeries.CategoryData = "A2:A11";

            // Optional: format the chart (titles, legend, etc.)
            chart.Title.Text = "Sales with Moving Average";
            chart.Legend.Position = LegendPositionType.Bottom;

            // Save the workbook (lifecycle rule: save)
            workbook.Save("MovingAverageChart.xlsx", SaveFormat.Xlsx);
        }
    }
}
