// Title: C# – Add a Moving‑Average Trendline to a Line Chart and Save Its Equation with Aspose.Cells
// Description: Loads an existing workbook, inserts sample X/Y data, creates a line chart, adds a 3‑point moving‑average trendline to the first series, enables equation display, writes a note about the equation to cell D1, and saves the file as an Excel workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | moving average trendline | line chart | Excel chart equation | trendline period | display equation | store equation in cell | chart customization | Aspose.Cells example
// Common Searches: Aspose.Cells add moving average trendline C# | display trendline equation Aspose.Cells .NET | set moving average period for chart series Aspose.Cells | write chart equation to worksheet cell Aspose.Cells | line chart trendline customization Aspose.Cells
// Developer Intent: Add a moving‑average trendline to a line chart, show its equation on the chart, and record a reference to that equation in a worksheet cell.
// Use Cases: Generate a line chart from worksheet data and smooth the series with a 3‑point moving average. | Configure the trendline to display its equation, assign a custom name, and apply a red color for visual emphasis. | Insert a textual note in a worksheet cell indicating that the equation is visible on the chart before saving the workbook.
// AI Prompts: Write C# code that adds a moving‑average trendline with a custom period to an Aspose.Cells line chart and enables equation display. | Show how to extract the trendline equation from an Aspose.Cells chart and write it into a specific worksheet cell. | Explain how to style a moving‑average trendline (color, name, period) using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

// Loads an existing workbook, inserts sample X/Y data, creates a line chart, adds a 3‑point moving‑average trendline to the first series, enables equation display, writes a note about the equation to cell D1, and saves the file as an Excel workbook using Aspose.Cells for .NET.
class MovingAverageTrendlineExample
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // Prepare sample data for the chart (if not already present)
        // -------------------------------------------------
        // X values
        sheet.Cells["A1"].PutValue("X");
        for (int i = 2; i <= 10; i++)
            sheet.Cells[$"A{i}"].PutValue(i - 1);

        // Y values
        sheet.Cells["B1"].PutValue("Y");
        for (int i = 2; i <= 10; i++)
            sheet.Cells[$"B{i}"].PutValue((i - 1) * 2 + (i % 3)); // some sample data

        // -------------------------------------------------
        // Add a line chart
        // -------------------------------------------------
        int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data source for the series
        chart.NSeries.Add("B2:B10", true);          // Y values
        chart.NSeries.CategoryData = "A2:A10";     // X values

        // -------------------------------------------------
        // Add a Moving Average trendline to the first series
        // -------------------------------------------------
        int trendlineIdx = chart.NSeries[0].TrendLines.Add(TrendlineType.MovingAverage);
        Trendline trendline = chart.NSeries[0].TrendLines[trendlineIdx];

        // Configure the trendline
        trendline.Period = 3;                       // 3‑point moving average
        trendline.Name = "3‑Period Moving Average";
        trendline.DisplayEquation = true;           // Show equation on the chart
        trendline.DisplayRSquared = false;          // R‑squared not needed for moving average
        trendline.Color = Color.Red;                // Optional visual styling

        // -------------------------------------------------
        // Store the equation text in a worksheet cell
        // (Aspose.Cells does not expose the equation string directly,
        //  so we store a placeholder indicating that the equation is displayed on the chart.)
        // -------------------------------------------------
        sheet.Cells["D1"].PutValue("Equation is displayed on the chart.");

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
