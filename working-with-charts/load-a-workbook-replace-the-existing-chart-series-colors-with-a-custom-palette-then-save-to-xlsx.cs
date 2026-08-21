// Title: How to Replace Excel Chart Series Colors with a Custom Palette Using Aspose.Cells for .NET (C#)
// Description: Loads an existing XLSX file, modifies the first five entries of the workbook palette with custom red, green, blue, orange, and purple colors via Workbook.ChangePalette, then iterates every worksheet and chart to apply a predefined ChartColorPaletteType (e.g., MonochromaticPalette5) to each series collection, and finally saves the workbook as a new file.
// Keywords: Aspose.Cells | Aspose.Cells for .NET | C# | ChangePalette | custom chart colors | Excel chart series color | ChartColorPaletteType | MonochromaticPalette5 | programmatic Excel styling | replace chart palette | load workbook | save workbook
// Common Searches: Aspose.Cells change chart series colors | set custom palette for Excel charts .NET | replace default Excel colors with custom palette using Aspose.Cells | apply monochromatic palette to all charts in a workbook | C# code to modify chart colors after loading an XLSX
// Developer Intent: Update the color palette of chart series in an existing workbook and persist the changes.
// Use Cases: Load an existing XLSX and redefine the first five palette entries with specific RGB values. | Loop through all worksheets and charts to assign a chosen ChartColorPaletteType to each series collection. | Save the modified workbook under a new filename while preserving all other data.
// AI Prompts: Show me C# code that uses Aspose.Cells.ChangePalette to set custom colors for the first five palette entries in a workbook. | Generate a script that iterates over every chart in a workbook and applies MonochromaticPalette5 to the series collection. | Explain how to combine custom palette updates with chart series color changes and correctly save the workbook using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an existing XLSX file, modifies the first five entries of the workbook palette with custom red, green, blue, orange, and purple colors via Workbook.ChangePalette, then iterates every worksheet and chart to apply a predefined ChartColorPaletteType (e.g., MonochromaticPalette5) to each series collection, and finally saves the workbook as a new file.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Define custom colors and replace the first few palette entries
        Color[] customColors = new Color[]
        {
            Color.FromArgb(255, 255, 0, 0),   // Red
            Color.FromArgb(255, 0, 255, 0),   // Green
            Color.FromArgb(255, 0, 0, 255),   // Blue
            Color.FromArgb(255, 255, 165, 0), // Orange
            Color.FromArgb(255, 128, 0, 128)  // Purple
        };

        for (int i = 0; i < customColors.Length; i++)
        {
            // Change palette entry at index i to the custom color
            workbook.ChangePalette(customColors[i], i);
        }

        // Iterate through all worksheets and their charts
        foreach (Worksheet ws in workbook.Worksheets)
        {
            foreach (Chart chart in ws.Charts)
            {
                // Apply a monochromatic palette to the series collection
                SeriesCollection seriesColl = chart.NSeries;
                // Example: use MonochromaticPalette5
                seriesColl.ChangeColors(ChartColorPaletteType.MonochromaticPalette5);
            }
        }

        // Save the modified workbook
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}
