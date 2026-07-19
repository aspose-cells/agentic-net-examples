// Title: Add a Moving‑Average Trendline with Equation to a Line Chart in Aspose.Cells for .NET (C#)
// Description: Load an existing workbook, create a line chart from column data, attach a moving‑average trendline to the first series, show its equation (and optional R‑squared), then save the updated file using Aspose.Cells for C#.
// Keywords: Aspose.Cells C# line chart trendline | moving average trendline Aspose.Cells | display trendline equation .NET | show R squared on chart Aspose | chart smoothing Aspose.Cells
// Common Searches: Aspose.Cells add moving average trendline C# | how to display trendline equation in Aspose.Cells chart | set period for moving average trendline Aspose.Cells | show R squared value on Aspose.Cells line chart | create line chart from workbook using Aspose.Cells
// Developer Intent: Insert a moving‑average trendline into a line chart, expose its formula (and optionally R‑squared), and persist the workbook.
// Use Cases: Sales dashboard: plot daily sales with a 3‑point moving average and show the smoothing formula on the chart. | Performance monitoring: generate a line chart that automatically includes the moving‑average equation for quick validation. | Financial reporting: embed both the moving‑average line and its R‑squared value to assess trend fit in quarterly reports.
// AI Prompts: Generate C# code using Aspose.Cells that loads a workbook, builds a line chart from ranges A2:A10 and B2:B10, adds a moving‑average trendline with a period of 5, displays the equation, and saves the result as output.xlsx. | Provide an Aspose.Cells example that creates a line chart, applies a moving‑average trendline, enables equation and R‑squared display, and writes the modified workbook to a new file.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

// Load an existing workbook, create a line chart from column data, attach a moving‑average trendline to the first series, show its equation (and optional R‑squared), then save the updated file using Aspose.Cells for C#.
class MovingAverageTrendlineExample
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a line chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
        Chart chart = worksheet.Charts[chartIndex];

        // Define the data range for the series (adjust the range as needed)
        // Here we assume data is in columns A (X values) and B (Y values) starting from row 2
        chart.NSeries.Add("B2:B10", true);
        chart.NSeries.CategoryData = "A2:A10";

        // Add a moving-average trendline to the first series
        int trendlineIdx = chart.NSeries[0].TrendLines.Add(TrendlineType.MovingAverage);
        Trendline trendline = chart.NSeries[0].TrendLines[trendlineIdx];

        // Set the period for the moving average (e.g., 3 points)
        trendline.Period = 3;

        // Display the equation on the chart
        trendline.DisplayEquation = true;

        // Optionally, display the R‑squared value as well
        trendline.DisplayRSquared = true;

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
