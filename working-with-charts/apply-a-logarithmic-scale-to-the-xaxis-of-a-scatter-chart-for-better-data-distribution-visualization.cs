// Title: C# Example: Scatter Chart with Logarithmic X‑Axis Using Aspose.Cells
// Description: This C# sample creates a workbook, fills columns A and B with exponential X values and linear Y values, adds a scatter chart, binds the series, enables a logarithmic scale on the X (Category) axis with base 2, sets axis titles, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells | C# | .NET | scatter chart | logarithmic X axis | log scale | CategoryAxis.IsLogarithmic | log base 2 | Excel chart example | GitHub | source code
// Common Searches: Aspose.Cells scatter chart log axis C# | how to set logarithmic X axis in Aspose.Cells | logarithmic scale for chart category axis .NET | C# example of scatter chart with log base 2 | Aspose.Cells chart axis properties
// Developer Intent: Add a logarithmic scale to the X‑axis of a scatter chart in a .NET workbook.
// Use Cases: Display exponential data on a readable X‑axis for scientific or financial analysis. | Compare linear Y trends against rapidly growing X values within the same chart. | Produce Excel reports that require base‑2 logarithmic scaling for measurement data.
// AI Prompts: Show how to change the X‑axis log base to e (natural logarithm) in the same chart. | Add a secondary Y‑axis with logarithmic scaling to the scatter chart. | Format the logarithmic X‑axis labels to show powers of two (e.g., 2⁰, 2¹, 2²).

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This C# sample creates a workbook, fills columns A and B with exponential X values and linear Y values, adds a scatter chart, binds the series, enables a logarithmic scale on the X (Category) axis with base 2, sets axis titles, and saves the file as an Excel workbook.
class LogScaleScatterChart
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add header cells
        sheet.Cells["A1"].PutValue("X");
        sheet.Cells["B1"].PutValue("Y");

        // Populate sample data: exponential X values and linear Y values
        for (int i = 2; i <= 10; i++)
        {
            // X values: 1, 2, 4, 8, ...
            sheet.Cells[$"A{i}"].PutValue(Math.Pow(2, i - 2));
            // Y values: 20, 30, 40, ...
            sheet.Cells[$"B{i}"].PutValue(i * 10);
        }

        // Add a scatter chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Scatter, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Bind Y values (required first argument) and then X values
        chart.NSeries.Add("B2:B10", true);
        chart.NSeries[0].XValues = "A2:A10";
        chart.NSeries[0].Name = "Sample Series";

        // Apply logarithmic scaling to the X‑axis (CategoryAxis)
        chart.CategoryAxis.IsLogarithmic = true;   // Axis.IsLogarithmic property
        chart.CategoryAxis.LogBase = 2;            // Axis.LogBase property (optional, base 2)

        // Optional: set axis titles for clarity
        chart.CategoryAxis.Title.Text = "Logarithmic X Axis (Base 2)";
        chart.ValueAxis.Title.Text = "Y Values";

        // Save the workbook with the chart
        workbook.Save("LogScaleScatterChart.xlsx");
    }
}
