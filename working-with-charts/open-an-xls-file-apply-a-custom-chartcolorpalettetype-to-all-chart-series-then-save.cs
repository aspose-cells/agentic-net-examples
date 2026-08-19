// Title: Apply a ChartColorPalette to All Charts in an XLS Workbook with Aspose.Cells for .NET
// Description: Load an existing XLS file, choose a ChartColorPaletteType (e.g., MonochromaticPalette1), iterate through each worksheet and chart, and use SeriesCollection.ChangeColors to set the palette for every series before saving the workbook.
// Keywords: Aspose.Cells C# chart color palette | ChartColorPaletteType example | SeriesCollection.ChangeColors | modify chart colors XLS | apply monochromatic palette Aspose | bulk chart styling .NET | iterate worksheets and charts Aspose.Cells | change Excel chart series colors C# | Excel chart theming programmatically
// Common Searches: how to change colors of all chart series in an XLS file using Aspose.Cells | apply a monochromatic palette to every chart in a workbook C# | Aspose.Cells iterate worksheets charts set series colors | SeriesCollection.ChangeColors usage example | bulk update chart colors Aspose.Cells .NET
// Developer Intent: Update every chart in an existing XLS workbook to use a single ChartColorPaletteType and save the modified file.
// Use Cases: Enforce corporate branding by standardizing chart colors across a workbook. | Improve print readability with a uniform monochromatic palette for all charts. | Allow end‑users to switch chart themes dynamically based on report settings.
// AI Prompts: Generate C# code that applies ChartColorPaletteType.MulticolorPalette2 to all charts in a workbook and saves it as XLSX using Aspose.Cells. | Explain the SeriesCollection.ChangeColors method and list all available ChartColorPaletteType values. | Create a snippet that selects different ChartColorPaletteType values for column and pie charts in the same workbook with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Load an existing XLS file, choose a ChartColorPaletteType (e.g., MonochromaticPalette1), iterate through each worksheet and chart, and use SeriesCollection.ChangeColors to set the palette for every series before saving the workbook.
class Program
{
    static void Main()
    {
        // Load the existing XLS workbook
        Workbook workbook = new Workbook("input.xls");

        // Choose a monochromatic palette to apply to chart series
        ChartColorPaletteType palette = ChartColorPaletteType.MonochromaticPalette1;

        // Iterate through all worksheets in the workbook
        foreach (Worksheet ws in workbook.Worksheets)
        {
            // Iterate through all charts on the worksheet
            foreach (Chart chart in ws.Charts)
            {
                // Get the series collection of the chart
                SeriesCollection seriesColl = chart.NSeries;

                // Apply the selected color palette to all series in the collection
                seriesColl.ChangeColors(palette);
            }
        }

        // Save the workbook with the updated chart colors
        workbook.Save("output.xls");
    }
}
