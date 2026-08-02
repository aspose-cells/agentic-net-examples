// Title: C# – Change First Chart Series Color to MonochromaticPalette6 with Aspose.Cells
// Description: Load an XLSX workbook, locate the first worksheet's first chart, apply the MonochromaticPalette6 palette to its series collection using Aspose.Cells, and save the modified file.
// Keywords: Aspose.Cells C# chart color palette | MonochromaticPalette6 | change chart series color | Excel chart styling Aspose | ChartColorPaletteType | C# modify Excel chart colors | Aspose.Cells ChangeColors
// Common Searches: Aspose.Cells change chart series color C# | apply MonochromaticPalette6 to Excel chart | how to recolor chart series with Aspose.Cells | C# code to set chart palette in XLSX | Aspose.Cells chart color customization example
// Developer Intent: Apply the MonochromaticPalette6 palette to the first chart's series in an existing XLSX file and save the result.
// Use Cases: Enforce a corporate monochrome theme across all charts in financial reports. | Update legacy Excel templates so every chart uses a consistent color scheme before distribution. | Automate batch recoloring of charts in generated workbooks to improve visual uniformity.
// AI Prompts: Write C# code that iterates through every chart in a workbook and sets a specified ChartColorPaletteType for each series. | Explain how to programmatically confirm that MonochromaticPalette6 was applied to a chart after saving the workbook. | Show how to assign different color palettes to charts based on their index or worksheet location using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Load an XLSX workbook, locate the first worksheet's first chart, apply the MonochromaticPalette6 palette to its series collection using Aspose.Cells, and save the modified file.
class ChangeChartSeriesColor
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (index 0)
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure there is at least one chart in the worksheet
        if (worksheet.Charts.Count == 0)
        {
            Console.WriteLine("No charts found in the worksheet.");
            return;
        }

        // Get the first chart
        Chart chart = worksheet.Charts[0];

        // Get the series collection of the chart
        SeriesCollection seriesColl = chart.NSeries;

        // Apply the MonochromaticPalette6 to the series collection
        seriesColl.ChangeColors(ChartColorPaletteType.MonochromaticPalette6);

        // Save the modified workbook
        workbook.Save("output.xlsx", SaveFormat.Xlsx);

        Console.WriteLine("Chart series color changed and workbook saved as output.xlsx");
    }
}
