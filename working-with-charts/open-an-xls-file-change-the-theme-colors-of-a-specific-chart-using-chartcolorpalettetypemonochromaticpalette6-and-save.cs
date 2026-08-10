// Title: Apply MonochromaticPalette6 to a chart in an existing XLS workbook with Aspose.Cells for .NET
// Description: Loads an XLS file, accesses the first worksheet, selects the first chart, changes its series colors to the MonochromaticPalette6 palette using Chart.NSeries.ChangeColors, and saves the workbook as a new file.
// Keywords: Aspose.Cells chart color palette | MonochromaticPalette6 C# | change chart colors XLS | Chart.NSeries.ChangeColors | Aspose.Cells .NET example
// Common Searches: how to set chart palette Aspose.Cells | apply monochrome colors to XLS chart | Aspose.Cells change chart theme colors | C# update chart colors in existing workbook
// Developer Intent: Modify the color scheme of a specific chart in an existing XLS workbook to use the MonochromaticPalette6 palette and persist the changes.
// Use Cases: Enforce a corporate monochrome style on legacy Excel reports before distribution. | Automate visual consistency for financial dashboards generated from older XLS files. | Integrate chart re‑theming into a migration pipeline that upgrades XLS workbooks to newer Office formats.
// AI Prompts: Generate C# code that opens an XLS file, finds a chart by its index, and applies ChartColorPaletteType.MonochromaticPalette6 to its series using Aspose.Cells. | Explain the effect of Chart.NSeries.ChangeColors in Aspose.Cells and how to verify the palette after saving the workbook. | Refactor the sample to accept a chart index and a palette enum as parameters, enabling dynamic color updates.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an XLS file, accesses the first worksheet, selects the first chart, changes its series colors to the MonochromaticPalette6 palette using Chart.NSeries.ChangeColors, and saves the workbook as a new file.
class ChangeChartThemeColors
{
    static void Main()
    {
        // Load the existing XLS workbook
        Workbook workbook = new Workbook("input.xls");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure there is at least one chart in the worksheet
        if (worksheet.Charts.Count == 0)
        {
            Console.WriteLine("No charts found in the worksheet.");
            return;
        }

        // Get the specific chart (e.g., the first chart)
        Chart chart = worksheet.Charts[0];

        // Change the color palette of the chart's series to MonochromaticPalette6
        chart.NSeries.ChangeColors(ChartColorPaletteType.MonochromaticPalette6);

        // Save the modified workbook
        workbook.Save("output.xls");
        Console.WriteLine("Chart theme colors updated and workbook saved as output.xls.");
    }
}
