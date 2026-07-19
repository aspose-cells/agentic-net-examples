// Title: C# Aspose.Cells: Scatter Chart with Logarithmic X‑Axis (Base 2)
// Description: Creates a workbook, fills columns with exponential X and linear Y values, inserts a scatter chart, binds the series, sets the X‑axis to a base‑2 logarithmic scale, adds an axis title, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells | C# | scatter chart | logarithmic X axis | log base 2 | CategoryAxis.IsLogarithmic | Excel chart programmatically | chart axis scaling | .NET Excel library
// Common Searches: Aspose.Cells scatter chart logarithmic X axis | C# create scatter chart with log scale using Aspose.Cells | set chart axis base to 2 Aspose.Cells | enable logarithmic axis in Aspose.Cells chart | example exponential X values scatter chart Aspose.Cells
// Developer Intent: Generate a scatter chart, bind custom X/Y ranges, apply a base‑2 logarithmic scale to the X‑axis, and export the workbook.
// Use Cases: Display data that spans several orders of magnitude on a clear, spaced‑out X‑axis. | Produce automated Excel reports where exponential trends need visual emphasis. | Build reusable chart templates that allow dynamic adjustment of the logarithmic base.
// AI Prompts: Write C# code with Aspose.Cells to create a scatter chart whose X‑axis uses a base‑2 logarithmic scale. | Explain the effect of CategoryAxis.IsLogarithmic and LogBase on chart rendering in Aspose.Cells. | Provide step‑by‑step instructions to bind X and Y ranges to a scatter series and enable logarithmic scaling on the X‑axis.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, fills columns with exponential X and linear Y values, inserts a scatter chart, binds the series, sets the X‑axis to a base‑2 logarithmic scale, adds an axis title, and saves the file as an Excel workbook.
class LogarithmicScatterChart
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data: exponential X values and linear Y values
        sheet.Cells["A1"].PutValue("X Values");
        sheet.Cells["B1"].PutValue("Y Values");
        for (int i = 2; i <= 10; i++)
        {
            sheet.Cells[$"A{i}"].PutValue(Math.Pow(2, i - 2)); // 1, 2, 4, 8, ...
            sheet.Cells[$"B{i}"].PutValue(i * 10);            // 20, 30, 40, ...
        }

        // Insert a scatter chart
        int chartIndex = sheet.Charts.Add(ChartType.Scatter, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Bind Y values and X values to the first series
        chart.NSeries.Add("B2:B10", true);   // Y values
        chart.NSeries[0].XValues = "A2:A10"; // X values
        chart.NSeries[0].Name = "Exponential Data";

        // Apply logarithmic scaling to the X‑axis (category axis)
        chart.CategoryAxis.IsLogarithmic = true; // enable logarithmic scale
        chart.CategoryAxis.LogBase = 2;          // set base to 2 for clearer spacing
        chart.CategoryAxis.Title.Text = "Logarithmic X Axis (Base 2)";

        // Save the workbook with the chart
        workbook.Save("LogarithmicScatterChart.xlsx");
    }
}
