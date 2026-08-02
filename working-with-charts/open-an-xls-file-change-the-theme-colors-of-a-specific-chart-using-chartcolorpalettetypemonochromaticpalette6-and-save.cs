// Title: C# – Change Chart Theme to MonochromaticPalette6 in an XLS File Using Aspose.Cells
// Description: Load an XLS workbook with Aspose.Cells, locate the first chart, set its series palette to ChartColorPaletteType.MonochromaticPalette6, and save the updated file.
// Keywords: Aspose.Cells | C# chart color palette | MonochromaticPalette6 | ChartColorPaletteType | modify XLS chart colors | ChangeColors method | programmatic chart styling
// Common Searches: Aspose.Cells set chart palette C# | apply MonochromaticPalette6 to XLS chart | change chart colors without Excel | C# change series colors in existing workbook | Aspose.Cells chart theme example
// Developer Intent: Programmatically apply the MonochromaticPalette6 palette to a chart’s series in an existing XLS workbook and persist the changes.
// Use Cases: Enforce a uniform monochrome look across all charts in quarterly reports before distribution. | Implement corporate branding by automatically assigning a predefined palette to charts generated from data exports. | Replace default Excel colors with a high‑contrast scheme to improve readability for printed or screen‑shared workbooks.
// AI Prompts: Generate C# code that opens an XLS file with Aspose.Cells, selects a specific chart, applies ChartColorPaletteType.MonochromaticPalette6 to its series, and saves the workbook. | Show an example that iterates through every chart in a workbook and sets each series collection to MonochromaticPalette6 using the ChangeColors method. | Write a reusable function that receives a workbook path and a chart index, changes the chart’s color palette to MonochromaticPalette6, and returns the path of the saved file.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Load an XLS workbook with Aspose.Cells, locate the first chart, set its series palette to ChartColorPaletteType.MonochromaticPalette6, and save the updated file.
class ChangeChartThemeColors
{
    static void Main()
    {
        // Path to the existing XLS file
        string inputPath = "input.xls";

        // Load the workbook
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure there is at least one chart in the worksheet
        if (worksheet.Charts.Count == 0)
        {
            Console.WriteLine("No charts found in the worksheet.");
            return;
        }

        // Get the first chart (replace index if a specific chart is required)
        Chart chart = worksheet.Charts[0];

        // Change the color palette of all series in the chart to MonochromaticPalette6
        SeriesCollection seriesColl = chart.NSeries;
        seriesColl.ChangeColors(ChartColorPaletteType.MonochromaticPalette6);

        // Save the modified workbook
        string outputPath = "output.xls";
        workbook.Save(outputPath);

        Console.WriteLine($"Chart theme colors updated and saved to '{outputPath}'.");
    }
}
