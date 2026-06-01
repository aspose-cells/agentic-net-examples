using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class TrendlineExample
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("InputWorkbook.xlsx");

        // Assume the first worksheet contains at least one chart
        Worksheet sheet = workbook.Worksheets[0];
        if (sheet.Charts.Count == 0)
        {
            Console.WriteLine("No charts found in the first worksheet.");
            return;
        }

        // Get the first chart in the worksheet
        Chart chart = sheet.Charts[0];

        // Ensure the chart has at least one series
        if (chart.NSeries.Count == 0)
        {
            Console.WriteLine("The chart does not contain any series.");
            return;
        }

        // Add a linear trendline to the first series
        int trendlineIndex = chart.NSeries[0].TrendLines.Add(TrendlineType.Linear);
        Trendline trendline = chart.NSeries[0].TrendLines[trendlineIndex];

        // Configure the trendline to display its equation on the chart
        trendline.DisplayEquation = true;
        trendline.DisplayRSquared = false; // optional: hide R‑squared

        // Log information about the added trendline
        Console.WriteLine("Trendline added to series 0:");
        Console.WriteLine($"- Type: {trendline.Type}");
        Console.WriteLine($"- DisplayEquation: {trendline.DisplayEquation}");
        // Note: Aspose.Cells does not expose the equation string directly.
        // The equation will be visible on the chart when opened in Excel.

        // Save the workbook (replace with your desired output path)
        workbook.Save("OutputWorkbook.xlsx");
    }
}