// Title: C# – Set CustomPalette2 for the Third Chart Series in an XLSX Workbook using Aspose.Cells
// Description: Loads an XLSX file, accesses the first worksheet and its first chart, verifies at least three series exist, changes the workbook palette entry at index 2 to a custom gray shade, applies that color to the third series with custom formatting, and saves the workbook as output.xlsx.
// Keywords: Aspose.Cells C# chart custom palette | ChartColorPaletteType.CustomPalette2 | change third series color Aspose.Cells | ChangePalette method .NET | assign custom color to chart series | modify chart series color programmatically
// Common Searches: How to apply CustomPalette2 to a specific chart series with Aspose.Cells | Set color of third series in Excel chart using C# Aspose.Cells | Change chart series palette entry in XLSX via Aspose.Cells | C# code to customize third series color in Excel chart
// Developer Intent: Apply a custom palette color (CustomPalette2) to the third series of the first chart in an XLSX workbook and save the changes.
// Use Cases: Brand a column chart by giving the third data series a custom gray tone. | Batch‑process multiple reports to ensure the third series uses a predefined palette entry. | Create a template workbook where the third series is pre‑colored, then populate data without losing the color.
// AI Prompts: Show C# code that sets ChartColorPaletteType.CustomPalette2 for the third series of a chart using Aspose.Cells. | Provide an Aspose.Cells example to change the third series color in an Excel chart to a custom palette entry and save the file.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an XLSX file, accesses the first worksheet and its first chart, verifies at least three series exist, changes the workbook palette entry at index 2 to a custom gray shade, applies that color to the third series with custom formatting, and saves the workbook as output.xlsx.
class AssignCustomPaletteToThirdSeries
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Assume the worksheet contains at least one chart
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

        // Define a custom color for the palette (this will act as "CustomPalette2")
        // Index 2 corresponds to the third entry in the 56‑color palette (0‑based)
        Color customPaletteColor = Color.FromArgb(255, 200, 200, 200);
        workbook.ChangePalette(customPaletteColor, 2);

        // Retrieve the color from the palette (now guaranteed to be present)
        Color paletteColor = workbook.Colors[2];

        // Assign the palette color to the third series (index 2)
        seriesColl[2].Area.ForegroundColor = paletteColor;
        seriesColl[2].Area.Formatting = FormattingType.Custom; // Ensure custom formatting is used

        // Save the modified workbook
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
