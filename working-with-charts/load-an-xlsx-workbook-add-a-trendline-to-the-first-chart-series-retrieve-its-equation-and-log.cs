// Title: Add a Linear Trendline to the First Chart Series in an XLSX Workbook with Aspose.Cells for .NET
// Description: Load an XLSX file using Aspose.Cells, locate the first worksheet and its first chart, attach a linear trendline to the chart's first series, enable equation display, read the intercept value, and save the updated workbook.
// Keywords: Aspose.Cells trendline | C# add chart trendline | Excel trendline equation Aspose | retrieve trendline intercept | save workbook Aspose.Cells | chart series trendline .NET | Aspose.Cells chart API | Linear trendline type | display trendline equation | Excel automation C#
// Common Searches: how to add a linear trendline to a chart using Aspose.Cells C# | Aspose.Cells retrieve trendline equation from Excel chart | C# get trendline intercept with Aspose.Cells | programmatically add regression line to Excel chart .NET | Aspose.Cells chart trendline example
// Developer Intent: Add a linear trendline to the first series of a chart, show its equation, capture the intercept, and write the changes back to the workbook.
// Use Cases: Generate monthly sales reports that automatically include a linear regression line and its equation on charts. | Extract the intercept of a trendline for downstream statistical analysis in a data‑processing pipeline. | Create a reusable Excel template that programmatically receives trendlines on all charts before distribution.
// AI Prompts: Write C# code with Aspose.Cells to add a polynomial trendline to the second series of a chart and display both the equation and R‑squared value. | Create a method that accepts a workbook path, adds a linear trendline to every series in every chart, logs each series' slope and intercept, and saves the file. | Show how to retrieve both the slope and intercept of a linear trendline from an Excel chart using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Load an XLSX file using Aspose.Cells, locate the first worksheet and its first chart, attach a linear trendline to the chart's first series, enable equation display, read the intercept value, and save the updated workbook.
class TrendlineExample
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (you can change the index as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure the worksheet contains at least one chart
        if (worksheet.Charts.Count == 0)
        {
            Console.WriteLine("No charts found in the worksheet.");
            return;
        }

        // Get the first chart in the worksheet
        Chart chart = worksheet.Charts[0];

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
        trendline.DisplayRSquared = false; // optional: hide R‑squared if not needed

        // Retrieve basic information about the trendline.
        // Aspose.Cells does not expose the full equation string directly,
        // but you can obtain key parameters such as the intercept.
        double intercept = trendline.Intercept;
        Console.WriteLine($"Trendline added. Type: {trendline.Type}, Intercept: {intercept}");

        // Save the modified workbook to a new file
        workbook.Save("output.xlsx");

        Console.WriteLine("Workbook saved as 'output.xlsx'.");
    }
}
