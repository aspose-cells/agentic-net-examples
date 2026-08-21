// Title: Aspose.Cells for .NET – Add a Moving‑Average Trendline with Equation to a Line Chart
// Description: Load an Excel workbook, create (or use) a line chart, attach a moving‑average trendline to the first series, set its period and custom name, enable the equation label, and save the file—all with Aspose.Cells C# API.
// Keywords: Aspose.Cells moving average trendline | display trendline equation .NET | C# line chart trendline Aspose | set trendline period Aspose.Cells | Excel chart equation Aspose.Cells | add trendline to chart programmatically
// Common Searches: Aspose.Cells add moving average trendline C# | show trendline equation in Excel chart using Aspose | set moving average period for chart series Aspose.Cells | create line chart with trendline Aspose.Cells .NET | how to display R‑squared value in Aspose.Cells chart
// Developer Intent: Insert a moving‑average trendline into a line chart and make its equation visible in an Excel workbook via Aspose.Cells for .NET.
// Use Cases: Automate financial dashboards by adding a 3‑period moving‑average trendline with a custom label to a sales line chart. | Enhance scientific reports with trendline equations for quick data interpretation without manual Excel editing. | Generate batch Excel files where each chart includes a configurable moving‑average trendline and its formula for downstream analysis.
// AI Prompts: Generate C# code with Aspose.Cells that loads a workbook, builds a line chart from A2:A10 and B2:B10, adds a 5‑period moving‑average trendline named "5‑Period MA", and displays its equation. | Explain how to retrieve the equation string of a moving‑average trendline after saving the workbook with Aspose.Cells. | Provide a step‑by‑step guide to add multiple moving‑average trendlines to different series in the same chart and show each equation using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

// Load an Excel workbook, create (or use) a line chart, attach a moving‑average trendline to the first series, set its period and custom name, enable the equation label, and save the file—all with Aspose.Cells C# API.
class AddMovingAverageTrendline
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (or any worksheet you need)
        Worksheet worksheet = workbook.Worksheets[0];

        // ------------------------------------------------------------
        // Create a line chart (if a chart already exists you can skip this)
        // ------------------------------------------------------------
        int chartIndex = worksheet.Charts.Add(ChartType.Line, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // ------------------------------------------------------------
        // Define the data range for the chart series
        // Adjust the ranges according to your worksheet data
        // ------------------------------------------------------------
        // Example: Y values in B2:B10, X (category) values in A2:A10
        chart.NSeries.Add("B2:B10", true);
        chart.NSeries.CategoryData = "A2:A10";

        // ------------------------------------------------------------
        // Add a Moving Average trendline to the first series
        // ------------------------------------------------------------
        int trendlineIndex = chart.NSeries[0].TrendLines.Add(TrendlineType.MovingAverage);
        Trendline trendline = chart.NSeries[0].TrendLines[trendlineIndex];

        // Set the period for the moving average (optional, default is 2)
        trendline.Period = 3;

        // Give the trendline a custom name (optional)
        trendline.Name = "3‑Period Moving Average";

        // Enable the display of the equation on the chart
        trendline.DisplayEquation = true;

        // (Optional) You can also display the R‑squared value
        // trendline.DisplayRSquared = true;

        // ------------------------------------------------------------
        // Save the workbook with the new chart and trendline
        // ------------------------------------------------------------
        workbook.Save("output.xlsx");
    }
}
