// Title: Apply a Custom ChartColorPaletteType to Every Chart in an XLS Workbook with Aspose.Cells (C#)
// Description: Loads an XLS file, loops through all worksheets and charts, sets a chosen ChartColorPaletteType for each series using NSeries.ChangeColors, and saves the workbook.
// Keywords: Aspose.Cells chart color palette | C# change chart series colors | ChartColorPaletteType example | iterate worksheets charts Aspose | modify XLS chart colors programmatically | NSeries.ChangeColors usage
// Common Searches: how to set chart color palette in Aspose.Cells C# | change all chart series colors in an XLS file | apply custom ChartColorPaletteType to workbook charts | Aspose.Cells iterate over charts and modify colors | save XLS after updating chart palettes
// Developer Intent: Programmatically assign a specific ChartColorPaletteType to every chart series in an existing XLS workbook and persist the changes.
// Use Cases: Enforce corporate branding by applying a uniform palette to all charts in financial reports. | Prepare presentation‑ready workbooks with consistent chart colors before distribution. | Automate compliance checks that require a predefined color scheme across generated XLS files.
// AI Prompts: Write C# code using Aspose.Cells to set ChartColorPaletteType.SpectrumPalette for all chart series in a workbook. | Show how to safely skip worksheets that contain no charts when applying NSeries.ChangeColors. | Explain how to let users select a ChartColorPaletteType at runtime and apply it to every chart in an XLS file.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an XLS file, loops through all worksheets and charts, sets a chosen ChartColorPaletteType for each series using NSeries.ChangeColors, and saves the workbook.
class Program
{
    static void Main()
    {
        // Load the existing XLS file
        Workbook workbook = new Workbook("input.xls");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet worksheet in workbook.Worksheets)
        {
            // Iterate through all charts on the current worksheet
            foreach (Chart chart in worksheet.Charts)
            {
                // Apply a monochromatic color palette to all series in the chart
                // You can change the enum value to use a different palette if needed
                chart.NSeries.ChangeColors(ChartColorPaletteType.MonochromaticPalette1);
            }
        }

        // Save the modified workbook to a new file
        workbook.Save("output.xls");
    }
}
