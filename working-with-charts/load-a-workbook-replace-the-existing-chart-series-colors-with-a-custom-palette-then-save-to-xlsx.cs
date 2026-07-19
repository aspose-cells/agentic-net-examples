// Title: Change Excel chart series colors using a custom palette with Aspose.Cells for .NET
// Description: Loads an existing XLSX workbook, accesses the first worksheet and its first chart, defines three RGB colors, overwrites the first three entries of the 56‑color workbook palette with those colors via ChangePalette, and saves the updated file. The chart automatically adopts the new palette indices.
// Keywords: Aspose.Cells chart color palette | C# change Excel chart series colors | custom workbook palette Aspose | replace chart colors programmatically | ChangePalette Aspose.Cells
// Common Searches: how to modify Excel chart colors with Aspose.Cells C# | replace default chart palette in an existing workbook | set custom RGB colors for chart series using Aspose | update Excel chart palette programmatically .NET | change chart series colors by editing workbook palette
// Developer Intent: Update the color scheme of chart series by editing the workbook's palette and save the modified workbook.
// Use Cases: Apply corporate brand colors to all charts in a template workbook automatically. | Generate recurring reports where each chart follows a predefined color scheme without manual editing. | Refresh existing chart series to reflect new palette entries after a bulk color update.
// AI Prompts: Show C# code that overwrites the first three palette entries with custom RGB values and makes existing charts use them in Aspose.Cells. | Provide an example of changing Excel chart series colors by updating the workbook palette and saving as XLSX. | Explain how to ensure charts refresh their series colors after calling ChangePalette with Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an existing XLSX workbook, accesses the first worksheet and its first chart, defines three RGB colors, overwrites the first three entries of the 56‑color workbook palette with those colors via ChangePalette, and saves the updated file. The chart automatically adopts the new palette indices.
class ReplaceChartSeriesColors
{
    static void Main()
    {
        // Load the existing workbook
        string inputFile = "input.xlsx";
        Workbook workbook = new Workbook(inputFile);

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure there is at least one chart in the worksheet
        if (worksheet.Charts.Count > 0)
        {
            // Get the first chart
            Chart chart = worksheet.Charts[0];

            // Define custom colors you want to use in the palette
            Color[] customPalette = new Color[]
            {
                Color.FromArgb(255, 255, 0, 0),   // Red
                Color.FromArgb(255, 0, 255, 0),   // Green
                Color.FromArgb(255, 0, 0, 255)    // Blue
            };

            // Replace the first three entries of the workbook palette with custom colors
            // (Excel palette has 56 entries, indices 0‑55)
            for (int i = 0; i < customPalette.Length; i++)
            {
                workbook.ChangePalette(customPalette[i], i);
            }

            // Optional: force the chart to refresh its series colors.
            // Changing the palette is enough for most cases because series use palette indices.
            // If needed, you can also apply a monochromatic palette to the series collection:
            // chart.NSeries.ChangeColors(ChartColorPaletteType.MonochromaticPalette1);
        }

        // Save the modified workbook to a new file
        string outputFile = "output.xlsx";
        workbook.Save(outputFile, SaveFormat.Xlsx);
    }
}
