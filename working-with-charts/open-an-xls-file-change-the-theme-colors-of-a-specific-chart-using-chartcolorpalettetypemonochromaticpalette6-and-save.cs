// Title: Change Chart Theme to MonochromaticPalette6 in an XLS Workbook (C# Aspose.Cells)
// Description: Loads an XLS file, accesses the first worksheet, finds the first chart, applies the MonochromaticPalette6 palette to its series via NSeries.ChangeColors, and saves the workbook.
// Keywords: Aspose.Cells | C# | .NET | XLS chart | chart color palette | MonochromaticPalette6 | ChartColorPaletteType | ChangeColors | NSeries | programmatic chart styling
// Common Searches: Aspose.Cells set chart palette MonochromaticPalette6 | C# change XLS chart colors programmatically | Apply monochrome theme to chart using Aspose.Cells | How to use NSeries.ChangeColors in Aspose.Cells
// Developer Intent: Apply a monochromatic color palette to a chart in an existing XLS workbook and save the result.
// Use Cases: Enforce corporate branding by standardizing chart colors across legacy XLS reports. | Automate theme updates for multiple charts during a reporting migration. | Replace default chart palettes with a specific monochrome scheme for visual consistency.
// AI Prompts: Generate C# code with Aspose.Cells that opens an XLS file, sets the first chart's series to ChartColorPaletteType.MonochromaticPalette6, and saves the workbook. | Show how to loop through all charts in a worksheet and apply MonochromaticPalette6 using Aspose.Cells. | Explain the NSeries.ChangeColors method and list alternative ChartColorPaletteType options for chart theming.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an XLS file, accesses the first worksheet, finds the first chart, applies the MonochromaticPalette6 palette to its series via NSeries.ChangeColors, and saves the workbook.
class ChangeChartThemeColors
{
    static void Main()
    {
        // Load the existing XLS workbook
        Workbook workbook = new Workbook("input.xls");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure there is at least one chart in the worksheet
        if (worksheet.Charts.Count > 0)
        {
            // Get the first chart (replace with specific index if a different chart is required)
            Chart chart = worksheet.Charts[0];

            // Change the color palette of the chart's series to MonochromaticPalette6 (Accent6 gradient)
            chart.NSeries.ChangeColors(ChartColorPaletteType.MonochromaticPalette6);
        }
        else
        {
            Console.WriteLine("No charts found in the worksheet.");
        }

        // Save the modified workbook
        workbook.Save("output.xls");
    }
}
