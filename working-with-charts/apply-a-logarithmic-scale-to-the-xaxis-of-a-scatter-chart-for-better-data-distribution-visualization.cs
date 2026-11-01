// Title: Apply a base‑10 logarithmic scale to the X‑axis of a scatter chart with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to generate a scatter chart, set the X‑axis to logarithmic scale with base 10, and save the workbook. | Show how to configure major and minor tick units for a logarithmic X‑axis in an Aspose.Cells scatter chart. | Demonstrate adding sample X/Y data, linking it to a scatter series, and exporting the chart as an .xlsx file with a log‑scaled X‑axis.
// Common Searches: Aspose.Cells C# scatter chart logarithmic X axis example | How to set log base 10 for X axis in Aspose.Cells chart | Configure major and minor tick spacing on a logarithmic axis using Aspose.Cells | Create Excel scatter plot with log‑scaled X axis in .NET | Programmatically apply logarithmic scaling to chart axis with Aspose.Cells
// Tags: Aspose.Cells set X axis logarithmic | C# scatter chart log base 10 | Aspose.Cells configure axis tick units | Create Excel scatter chart with log axis | Aspose.Cells chart scaling .NET

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// The sample creates a new workbook, fills columns A and B with X/Y values, adds a scatter chart, links the series to the data, enables a base‑10 logarithmic scale on the X (category) axis, adjusts major and minor tick units, and saves the result as ScatterLogScale.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Access the first worksheet
            Worksheet ws = wb.Worksheets[0];

            // Populate sample data for the scatter chart (X in column A, Y in column B)
            ws.Cells["A1"].PutValue("X");
            ws.Cells["B1"].PutValue("Y");
            double[] xValues = { 1, 10, 100, 1000, 10000 };
            double[] yValues = { 2, 20, 200, 2000, 20000 };
            for (int i = 0; i < xValues.Length; i++)
            {
                ws.Cells[i + 1, 0].PutValue(xValues[i]); // Column A
                ws.Cells[i + 1, 1].PutValue(yValues[i]); // Column B
            }

            // Add a scatter chart to the worksheet
            int chartIndex = ws.Charts.Add(ChartType.Scatter, 5, 0, 25, 10);
            Chart chart = ws.Charts[chartIndex];

            // Add a series: Y values from B2:B6, X values from A2:A6
            int seriesIndex = chart.NSeries.Add("B2:B6", true);
            chart.NSeries[seriesIndex].XValues = "A2:A6";

            // Apply logarithmic scale to the X‑axis (Category axis for XY scatter)
            Axis xAxis = chart.CategoryAxis;
            xAxis.IsLogarithmic = true;   // Enable logarithmic scaling
            xAxis.LogBase = 10;           // Use base‑10 logarithm

            // Optional: adjust major/minor units for better tick spacing
            xAxis.MajorUnit = 1;
            xAxis.MinorUnit = 0.1;

            // Save the workbook with the chart
            wb.Save("ScatterLogScale.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
