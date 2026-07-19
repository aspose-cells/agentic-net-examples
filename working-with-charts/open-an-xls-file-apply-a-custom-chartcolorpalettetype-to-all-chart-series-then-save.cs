// Title: Apply a Monochromatic ChartColorPalette to All Charts in an XLS Workbook (C# Aspose.Cells)
// Description: Loads an XLS workbook, selects a monochromatic ChartColorPaletteType, iterates through each worksheet and chart, applies the palette to all series with SeriesCollection.ChangeColors, and saves the updated file.
// Keywords: Aspose.Cells | C# | ChartColorPaletteType | ChangeColors | apply chart palette | XLS chart colors | chart series styling | .NET chart customization | monochromatic palette | workbook chart iteration
// Common Searches: Aspose.Cells change chart colors C# | apply ChartColorPaletteType to all charts in a workbook | set monochrome palette for XLS charts using Aspose.Cells | iterate through charts in a workbook Aspose.Cells | SeriesCollection.ChangeColors example
// Developer Intent: Load an existing XLS file, apply a chosen ChartColorPaletteType to every chart series, and save the modified workbook.
// Use Cases: Ensure a consistent visual theme across financial dashboards by applying a single monochromatic palette to all charts. | Prepare presentation‑ready workbooks with uniform chart colors before sharing with stakeholders. | Automate corporate rebranding of legacy XLS reports by updating every chart to the new brand color scheme.
// AI Prompts: Generate C# code with Aspose.Cells that applies ChartColorPaletteType.EarthTonesPalette to all charts in a workbook and saves it as .xlsx. | Explain how SeriesCollection.ChangeColors works with different chart types and note any limitations or required chart formats. | Create robust error‑handling logic for scenarios where a workbook contains no charts, uses an unsupported format, or the palette cannot be applied.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an XLS workbook, selects a monochromatic ChartColorPaletteType, iterates through each worksheet and chart, applies the palette to all series with SeriesCollection.ChangeColors, and saves the updated file.
class ApplyChartPalette
{
    static void Main()
    {
        // Load the existing XLS workbook
        Workbook workbook = new Workbook("input.xls");

        // Define the desired monochromatic palette
        ChartColorPaletteType palette = ChartColorPaletteType.MonochromaticPalette1;

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all charts in the worksheet
            foreach (Chart chart in sheet.Charts)
            {
                // Get the series collection of the chart
                SeriesCollection seriesColl = chart.NSeries;

                // Apply the selected color palette to all series
                seriesColl.ChangeColors(palette);
            }
        }

        // Save the modified workbook
        workbook.Save("output.xls");
    }
}
