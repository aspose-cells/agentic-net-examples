// Title: Set Logarithmic Scale on Primary Y‑Axis of a Column Chart with Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, fills cells A1:B4 with values ranging from 10 to 1,000,000, adds a column chart, and configures the primary Y‑axis to use a logarithmic scale (base 10, min 1, max 10,000,000) before saving the file.
// Keywords: Aspose.Cells | C# | .NET chart logarithmic axis | primary Y axis log scale | column chart scaling | log base setting | axis min max Aspose.Cells | value axis logarithmic
// Common Searches: Aspose.Cells set Y axis to logarithmic scale C# | how to enable log scale on chart axis using Aspose.Cells | logarithmic value axis example Aspose.Cells .NET | column chart with large value range Aspose.Cells | configure chart axis limits Aspose.Cells C#
// Developer Intent: Apply a logarithmic scale to the primary Y‑axis of a chart so that values spanning multiple orders of magnitude are displayed clearly.
// Use Cases: Financial dashboards where revenues range from thousands to billions. | Scientific plots showing exponential growth or decay. | Performance metrics that include both low‑level and high‑level measurements on a single chart.
// AI Prompts: Show how to set a logarithmic Y‑axis with base 2 in Aspose.Cells for .NET. | Explain how to customize tick marks and label formatting after enabling a log scale on a chart axis. | Provide code that reads existing worksheet data and applies logarithmic scaling to the chart's value axis dynamically.

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// Creates a new workbook, fills cells A1:B4 with values ranging from 10 to 1,000,000, adds a column chart, and configures the primary Y‑axis to use a logarithmic scale (base 10, min 1, max 10,000,000) before saving the file.
class LogarithmicAxisDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data with a wide range of values
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("Low");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("Medium");
            worksheet.Cells["B3"].PutValue(1_000);
            worksheet.Cells["A4"].PutValue("High");
            worksheet.Cells["B4"].PutValue(1_000_000);

            // Add a column chart and bind the data
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = worksheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Apply logarithmic scaling to the primary Y axis
            chart.ValueAxis.IsLogarithmic = true;   // Enable logarithmic scale
            chart.ValueAxis.LogBase = 10;           // Set logarithmic base (default is 10)
            chart.ValueAxis.MinValue = 1;           // Minimum value (must be > 0)
            chart.ValueAxis.MaxValue = 10_000_000; // Optional maximum value

            // Save the workbook
            workbook.Save("LogarithmicAxisDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

class Program
{
    static void Main(string[] args)
    {
        LogarithmicAxisDemo.Run();
    }
}
