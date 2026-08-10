// Title: C# – Apply a Uniform Chart Color Palette to All Charts and Save as XLS with Aspose.Cells
// Description: Load an XLSX workbook, loop through each worksheet and chart, apply a chosen ChartColorPaletteType to all series via Chart.NSeries.ChangeColors, then export the file as an Excel 97‑2003 XLS using XlsSaveOptions.MatchColor to map colors to the 56‑color palette.
// Keywords: Aspose.Cells | C# chart color palette | ChartColorPaletteType | Chart.NSeries.ChangeColors | save as XLS | XlsSaveOptions MatchColor | uniform chart styling | Excel 97-2003 export | batch chart formatting | legacy Excel compatibility
// Common Searches: Aspose.Cells set same chart colors for all worksheets | How to change chart palette in C# Aspose.Cells | Save XLSX with charts to XLS preserving colors | Chart.NSeries.ChangeColors example | Map chart colors to 56‑color palette Aspose
// Developer Intent: Apply a single color palette to every chart in a workbook and export the workbook to the legacy XLS format while keeping chart colors consistent.
// Use Cases: Standardize chart appearance across a multi‑sheet report before distributing to clients using older Excel versions. | Batch‑process corporate workbooks to enforce a company‑wide chart color scheme. | Convert modern XLSX files with charts to XLS for compatibility with legacy systems, ensuring colors map to the 56‑color palette.
// AI Prompts: Generate C# code that loads an XLSX file, applies a specific ChartColorPaletteType to all charts using Aspose.Cells, and saves the result as XLS with MatchColor enabled. | Explain how Chart.NSeries.ChangeColors works in Aspose.Cells and advise on selecting an appropriate ChartColorPaletteType for uniform styling. | Provide a step‑by‑step tutorial for converting an XLSX workbook containing charts to XLS while preserving chart colors with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Load an XLSX workbook, loop through each worksheet and chart, apply a chosen ChartColorPaletteType to all series via Chart.NSeries.ChangeColors, then export the file as an Excel 97‑2003 XLS using XlsSaveOptions.MatchColor to map colors to the 56‑color palette.
class ApplyUniformChartPalette
{
    static void Main()
    {
        // Load the existing workbook (replace the path with your source file)
        Workbook workbook = new Workbook("input.xlsx");

        // Choose a chart color palette type (using the first enum value as a generic example)
        ChartColorPaletteType paletteType = (ChartColorPaletteType)0;

        // Apply the selected palette to every chart in every worksheet
        foreach (Worksheet ws in workbook.Worksheets)
        {
            foreach (Chart chart in ws.Charts)
            {
                // Change the colors of all series in the chart
                chart.NSeries.ChangeColors(paletteType);
            }
        }

        // Configure save options for the Excel 97‑2003 format
        XlsSaveOptions saveOptions = new XlsSaveOptions
        {
            MatchColor = true // map colors to the 56‑color palette
        };

        // Save the workbook as an XLS file
        workbook.Save("output.xls", saveOptions);
    }
}
