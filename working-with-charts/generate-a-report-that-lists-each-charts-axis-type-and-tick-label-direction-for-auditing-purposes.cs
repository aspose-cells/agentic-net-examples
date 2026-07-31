// Title: C# console report of chart axis types and tick‑label directions with Aspose.Cells
// Description: Loads an Excel workbook, iterates every worksheet and chart, calculates each chart layout, and prints to the console the axis scope (primary/secondary), axis type (Category, Value, Series) and tick‑label text direction. Ideal for .NET developers needing a quick audit of chart axis settings.
// Keywords: Aspose.Cells chart axis direction | C# list chart axes tick label | retrieve axis type Aspose.Cells | audit Excel chart axes .NET | chart axis tick label direction C# | Aspose.Cells console report | Excel chart axis enumeration
// Common Searches: how to get tick label direction of chart axes using Aspose.Cells C# | list primary and secondary axes for all charts in a workbook | Aspose.Cells enumerate chart axes properties | C# code to output chart axis type and label direction | Aspose.Cells chart axis audit example
// Developer Intent: Generate a detailed console report that shows each chart’s axis type and its tick‑label direction across all worksheets in an Excel file.
// Use Cases: Create an audit log of chart axis settings before publishing a workbook to guarantee visual consistency. | Validate that tick‑label directions follow corporate style guidelines in financial or marketing reports. | Troubleshoot unexpected axis formatting in automated chart generation pipelines.
// AI Prompts: Write a method that returns a List<ChartAxisInfo> with worksheet name, chart name, axis scope, axis type, and tick‑label direction. | Modify the sample to export the axis report to CSV or JSON instead of writing to the console. | Add error handling for missing secondary axes and produce a summary of primary vs. secondary axes found.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an Excel workbook, iterates every worksheet and chart, calculates each chart layout, and prints to the console the axis scope (primary/secondary), axis type (Category, Value, Series) and tick‑label text direction. Ideal for .NET developers needing a quick audit of chart axis settings.
class ChartAxisAudit
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook that contains charts.
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all charts on the worksheet.
                foreach (Chart chart in sheet.Charts)
                {
                    // Ensure the chart layout is calculated before accessing axis properties.
                    chart.Calculate();

                    // Local helper to output information for a given axis.
                    void ReportAxis(Axis axis, AxisType type, bool isPrimary)
                    {
                        if (axis == null) return; // Safety check.

                        // Get the text direction of the tick labels.
                        ChartTextDirectionType direction = axis.TickLabels.DirectionType;

                        // Build a readable description.
                        string axisScope = isPrimary ? "Primary" : "Secondary";
                        Console.WriteLine($"Worksheet: {sheet.Name}");
                        Console.WriteLine($"Chart: {(!string.IsNullOrEmpty(chart.Name) ? chart.Name : "Unnamed Chart")}");
                        Console.WriteLine($"  {axisScope} {type} Axis:");
                        Console.WriteLine($"    TickLabel Direction: {direction}");
                    }

                    // Primary Category Axis
                    if (chart.HasAxis(AxisType.Category, true))
                    {
                        ReportAxis(chart.CategoryAxis, AxisType.Category, true);
                    }

                    // Secondary Category Axis
                    if (chart.HasAxis(AxisType.Category, false))
                    {
                        ReportAxis(chart.SecondCategoryAxis, AxisType.Category, false);
                    }

                    // Primary Value Axis
                    if (chart.HasAxis(AxisType.Value, true))
                    {
                        ReportAxis(chart.ValueAxis, AxisType.Value, true);
                    }

                    // Secondary Value Axis
                    if (chart.HasAxis(AxisType.Value, false))
                    {
                        ReportAxis(chart.SecondValueAxis, AxisType.Value, false);
                    }

                    // Primary Series Axis (if applicable)
                    if (chart.HasAxis(AxisType.Series, true))
                    {
                        ReportAxis(chart.SeriesAxis, AxisType.Series, true);
                    }

                    // Secondary Series Axis (if applicable)
                    if (chart.HasAxis(AxisType.Series, false) && chart.SeriesAxis != null)
                    {
                        // Aspose.Cells does not have a distinct property for secondary series axis,
                        // so we reuse the same SeriesAxis reference when HasAxis returns true for false.
                        ReportAxis(chart.SeriesAxis, AxisType.Series, false);
                    }

                    Console.WriteLine(); // Blank line between charts
                }
            }

            // Optionally, save the workbook after any modifications.
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
