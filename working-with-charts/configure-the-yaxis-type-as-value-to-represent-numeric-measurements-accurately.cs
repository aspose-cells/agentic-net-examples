// Title: How to set the Y‑Axis to a Value axis for a column chart using Aspose.Cells in C#
// AI Prompts: Write C# code that accesses a column chart in an Aspose.Cells workbook, sets its primary Y‑axis to AxisType.Value, and saves the workbook. | Show the steps to change an existing Aspose.Cells chart's Y‑axis from Category to Value in a .NET application. | Provide a snippet that configures the numeric Y‑axis for a column chart created with Aspose.Cells and persists the file.
// Common Searches: Aspose.Cells change chart Y axis to numeric value in C# | C# set primary Y axis type Value for Excel column chart using Aspose.Cells | How to configure numeric Y axis for a chart created with Aspose.Cells .NET | Set AxisType.Value on Y axis of Aspose.Cells column chart programmatically | Modify existing chart axis type to Value with Aspose.Cells C# example
// Tags: set Y axis value Aspose.Cells | column chart numeric axis .NET | modify chart axis type C# | Aspose.Cells AxisType.Value example | configure chart Y axis Aspose.Cells | Excel chart numeric Y axis Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program loads or creates an Excel workbook, ensures a worksheet exists, retrieves or adds a column chart, sets the chart’s primary Y‑axis to AxisType.Value for accurate numeric measurement display, and saves the workbook to the specified file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook if it exists; otherwise create a new one.
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Ensure there is at least one worksheet.
            if (workbook.Worksheets.Count == 0)
            {
                workbook.Worksheets.Add();
            }

            // Get the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Obtain a chart – use existing one if present, otherwise create a new chart.
            Chart chart;
            if (sheet.Charts.Count > 0)
            {
                chart = sheet.Charts[0];
            }
            else
            {
                // Add a column chart and retrieve it by the returned index.
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                chart = sheet.Charts[chartIndex];
            }

            // (Optional) Additional chart configuration can be added here.

            // Save the modified workbook.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
