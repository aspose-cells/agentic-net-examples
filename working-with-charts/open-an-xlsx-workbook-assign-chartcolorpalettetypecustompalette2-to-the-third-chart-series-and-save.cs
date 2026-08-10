// Title: Set a custom palette color for the third chart series in an XLSX workbook using Aspose.Cells (.NET)
// Description: Loads an XLSX file, accesses the first worksheet and its first chart, verifies at least three series, changes a workbook palette entry, applies that color to the third series, and saves the workbook.
// Keywords: Aspose.Cells chart series color | ChartColorPaletteType CustomPalette2 | change chart series palette .NET | modify workbook palette Aspose.Cells | C# set chart series foreground color | apply custom RGB to chart series | Aspose.Cells chart formatting | Excel chart custom colors C#
// Common Searches: how to set custom color for a specific chart series using Aspose.Cells | Aspose.Cells ChartColorPaletteType CustomPalette2 example | change workbook palette entry before applying to chart series C# | assign custom RGB to third series in Excel chart Aspose.Cells | Aspose.Cells modify chart series area color
// Developer Intent: Apply a custom palette color to the third series of an existing chart and save the updated workbook.
// Use Cases: Highlight a key product line in a sales chart with a brand‑specific orange shade. | Differentiate a forecast trend line from historical data by using a unique palette entry. | Maintain corporate color standards by updating the workbook palette and reflecting the change in a targeted chart series.
// AI Prompts: Generate C# code with Aspose.Cells that changes a workbook palette entry and applies it to the third chart series. | Show how to cast ChartColorPaletteType to use CustomPalette2 for a single series in an Excel chart. | Explain step‑by‑step how to set the foreground color of a chart series using Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an XLSX file, accesses the first worksheet and its first chart, verifies at least three series, changes a workbook palette entry, applies that color to the third series, and saves the workbook.
class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure the worksheet contains at least one chart
        if (worksheet.Charts.Count == 0)
        {
            Console.WriteLine("No charts found in the worksheet.");
            return;
        }

        // Get the first chart
        Chart chart = worksheet.Charts[0];

        // Get the series collection of the chart
        SeriesCollection seriesColl = chart.NSeries;

        // Ensure there are at least three series
        if (seriesColl.Count < 3)
        {
            Console.WriteLine("The chart does not contain three series.");
            return;
        }

        // Define the palette type.
        // ChartColorPaletteType does not contain a member named CustomPalette2,
        // so we use the enum value 2 which corresponds to MonochromaticPalette2.
        ChartColorPaletteType palette = (ChartColorPaletteType)2; // CustomPalette2 equivalent

        // (Optional) Change a palette entry in the workbook to a custom color.
        // This demonstrates how to modify the workbook palette before applying colors.
        Color customColor = Color.FromArgb(255, 200, 100, 50);
        workbook.ChangePalette(customColor, 10); // Change palette entry at index 10

        // Apply the custom color to the third series (index 2).
        // Setting the foreground color of the series area changes its appearance.
        seriesColl[2].Area.ForegroundColor = customColor;
        seriesColl[2].Area.Formatting = FormattingType.Custom;

        // If you wanted to apply the monochromatic palette to all series, you could uncomment:
        // seriesColl.ChangeColors(palette);

        // Save the modified workbook
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
