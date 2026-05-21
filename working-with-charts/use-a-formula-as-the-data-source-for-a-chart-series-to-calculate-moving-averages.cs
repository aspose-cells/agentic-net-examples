using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class MovingAverageChart
{
    static void Main()
    {
        // Create a new workbook (lifecycle create)
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        Cells cells = ws.Cells;

        // Populate sample data in column B (B2:B11)
        for (int i = 0; i < 10; i++)
        {
            cells[i + 1, 1].PutValue(i + 1); // values 1..10
        }

        // Define moving‑average period
        int period = 3;

        // Insert moving‑average formulas in column C (C2:C11)
        for (int row = 1; row <= 10; row++)
        {
            // Determine the start row for the averaging window
            int startRow = Math.Max(1, row - period + 1);
            // Build the range string for the AVERAGE function, e.g., B2:B4
            string range = $"B{startRow + 1}:B{row + 1}";
            // Set the formula in the cell
            cells[row, 2].Formula = $"=AVERAGE({range})";
        }

        // Calculate all formulas so that moving‑average values are materialized
        wb.CalculateFormula();

        // Add a line chart to the worksheet
        int chartIndex = ws.Charts.Add(ChartType.Line, 5, 0, 20, 12);
        Chart chart = ws.Charts[chartIndex];

        // Use column A as the category (X) axis – fill it with row numbers
        for (int i = 0; i < 10; i++)
        {
            cells[i + 1, 0].PutValue(i + 1);
        }

        // Add the moving‑average series (C2:C11) and set its categories
        chart.NSeries.Add("C2:C11", true);
        chart.NSeries.CategoryData = "A2:A11";

        // Optional: add the original data series for comparison
        chart.NSeries.Add("B2:B11", true);
        chart.NSeries[1].Name = "Original";

        // Recalculate chart layout (optional but recommended)
        chart.Calculate();

        // Save the workbook (lifecycle save)
        wb.Save("MovingAverageChart.xlsx");
    }
}