// Title: Add a Linear Trendline and Show Its Equation in an Excel Chart with Aspose.Cells for .NET
// Description: Loads an XLSX workbook, verifies a chart exists, adds a linear trendline to the first series, enables equation display, logs the operation, and saves the updated file.
// Keywords: Aspose.Cells | C# | .NET | Excel chart trendline | add linear trendline | display trendline equation | chart series modification | programmatic Excel chart | workbook automation | trendline API
// Common Searches: Aspose.Cells add linear trendline C# | show trendline equation in Excel chart using Aspose.Cells | how to add trendline to first chart series Aspose.Cells .NET | retrieve trendline equation Aspose.Cells | programmatically modify Excel chart trendline
// Developer Intent: Programmatically insert a linear trendline into the first series of the first chart in an XLSX workbook and turn on equation display.
// Use Cases: Automatically attach a linear trendline to sales charts before generating quarterly financial reports. | Batch‑process multiple workbooks to add forecasting trendlines to their primary charts for distribution to stakeholders. | Enhance a data‑visualization dashboard by showing trendline equations directly on Excel charts for end‑user analysis.
// AI Prompts: Generate C# code with Aspose.Cells to add a polynomial trendline to the second series of a chart and display its equation. | Create a snippet that checks existing trendlines on a chart and updates the DisplayEquation property for each one. | Explain how to extract the equation string from a Trendline object after saving the workbook with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an XLSX workbook, verifies a chart exists, adds a linear trendline to the first series, enables equation display, logs the operation, and saves the updated file.
class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure the worksheet contains at least one chart
        if (worksheet.Charts.Count == 0)
        {
            Console.WriteLine("No charts found in the worksheet.");
            return;
        }

        // Get the first chart in the worksheet
        Chart chart = worksheet.Charts[0];

        // Add a linear trendline to the first series of the chart
        int trendlineIndex = chart.NSeries[0].TrendLines.Add(TrendlineType.Linear);
        Trendline trendline = chart.NSeries[0].TrendLines[trendlineIndex];

        // Enable the display of the equation on the chart
        trendline.DisplayEquation = true;

        // Log that the trendline has been added and its equation is set to be displayed
        Console.WriteLine($"Trendline added to series 0. DisplayEquation = {trendline.DisplayEquation}");

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
