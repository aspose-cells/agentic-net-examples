// Title: Hide PivotChart Legend and Drop Zones in an Existing Workbook with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, locates PivotCharts on the first worksheet, disables their drop‑zone controls via PivotOptions.DropZonesVisible = false, hides the chart legend with ShowLegend = false, and saves the updated file.
// Keywords: Aspose.Cells | C# | PivotChart | hide legend | drop zones | PivotOptions | DropZonesVisible | chart.ShowLegend | Excel chart formatting | programmatic Excel manipulation
// Common Searches: Aspose.Cells hide pivot chart legend C# | remove pivot chart drop zones .NET | set DropZonesVisible false Aspose | chart.ShowLegend false Aspose.Cells example | programmatically hide legend in PivotChart
// Developer Intent: Programmatically suppress the legend and pivot UI (drop zones) of PivotCharts in an existing Excel workbook using Aspose.Cells for .NET.
// Use Cases: Prepare a clean report workbook where PivotChart legends are removed for a professional look. | Create dashboards that prevent end‑users from altering pivot controls while keeping the visual layout tidy. | Export PivotCharts to PDF or images without legend clutter by disabling both the legend and drop zones.
// AI Prompts: Generate C# code with Aspose.Cells that hides the legend of a PivotChart and disables its drop zones in an existing workbook. | Show how to iterate through all charts in a worksheet and apply PivotOptions.DropZonesVisible = false only to charts linked to a PivotTable. | Explain how to verify that the legend and pivot UI elements are hidden after saving the workbook with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an Excel workbook, locates PivotCharts on the first worksheet, disables their drop‑zone controls via PivotOptions.DropZonesVisible = false, hides the chart legend with ShowLegend = false, and saves the updated file.
class HidePivotChartLegend
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Iterate through charts to find a PivotChart
        foreach (Chart chart in worksheet.Charts)
        {
            // Ensure the chart is linked to a PivotTable (PivotSource is not empty)
            if (!string.IsNullOrEmpty(chart.PivotSource))
            {
                // Access the PivotOptions of the chart
                PivotOptions pivotOptions = chart.PivotOptions;

                // Hide pivot controls (including legend-like elements) on the chart
                // Setting DropZonesVisible to false removes the pivot UI elements.
                pivotOptions.DropZonesVisible = false;

                // If you also want to hide the standard legend, you can set ShowLegend to false
                // (not part of PivotOptions, but commonly used for charts)
                chart.ShowLegend = false;
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
