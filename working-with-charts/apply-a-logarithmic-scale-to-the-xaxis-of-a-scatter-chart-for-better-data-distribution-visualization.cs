// Title: Create a Scatter Chart with a Logarithmic X‑Axis in Aspose.Cells for .NET (C#)
// Description: Shows how to build an Excel workbook, populate exponential X values and linear Y values, add a scatter chart, bind the series, and enable a logarithmic X‑axis (base 2) using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET charting | scatter chart | logarithmic X axis | CategoryAxis.IsLogarithmic | log base | Excel automation | data visualization | chart scaling
// Common Searches: Aspose.Cells set X axis logarithmic C# | logarithmic scatter chart example Aspose.Cells | change log base in Aspose.Cells chart | C# create scatter chart with log scale | CategoryAxis.IsLogarithmic property usage
// Developer Intent: Apply a logarithmic scale to the X‑axis of a scatter chart in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Display exponential or multi‑order‑of‑magnitude data where a log‑scaled X‑axis improves readability. | Generate scientific, engineering, or financial reports that require precise trend analysis on a log scale. | Programmatically adjust the log base (e.g., from 2 to 10) to match different domain conventions.
// AI Prompts: Provide C# code to set the X‑axis log base to 10 for an Aspose.Cells scatter chart. | Show how to toggle between linear and logarithmic scaling on an existing Aspose.Cells chart axis. | Explain how to customize X‑axis label formatting after enabling logarithmic scaling in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to build an Excel workbook, populate exponential X values and linear Y values, add a scatter chart, bind the series, and enable a logarithmic X‑axis (base 2) using Aspose.Cells for .NET.
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
            sheet.Cells[$"B{i}"].PutValue(i * 10);
        }

        // Add a scatter chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Scatter, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Add a series and bind Y values; then set X values separately
        chart.NSeries.Add("B2:B10", true);
        chart.NSeries[0].XValues = "A2:A10";
        chart.NSeries[0].Name = "Sample Series";

        // Configure the X‑axis (CategoryAxis) to use a logarithmic scale
        chart.CategoryAxis.IsLogarithmic = true;
        chart.CategoryAxis.LogBase = 2; // Set logarithmic base to 2 (can be changed)

        // Optional: add titles for clarity
        chart.Title.Text = "Scatter Chart with Logarithmic X‑Axis";
        chart.CategoryAxis.Title.Text = "Log Scale (Base 2)";
        chart.ValueAxis.Title.Text = "Y Value";

        // Save the workbook with the chart
        workbook.Save("LogarithmicScatterChart.xlsx");
    }
}
