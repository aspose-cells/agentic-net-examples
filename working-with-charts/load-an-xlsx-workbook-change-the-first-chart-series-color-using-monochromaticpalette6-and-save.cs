// Title: Change First Chart Series Color to MonochromaticPalette6 Using Aspose.Cells for .NET
// Description: Load an XLSX workbook, locate the first chart, apply the MonochromaticPalette6 palette to its series collection, and save the updated file with Aspose.Cells C# API.
// Keywords: Aspose.Cells chart color palette | MonochromaticPalette6 C# | change Excel chart series color | apply palette to chart Aspose | modify chart colors programmatically
// Common Searches: Aspose.Cells change chart series color C# | apply MonochromaticPalette6 to Excel chart | set chart palette with Aspose.Cells .NET | how to recolor chart series in XLSX using C#
// Developer Intent: Apply the MonochromaticPalette6 color scheme to the series of the first chart in an existing XLSX workbook and save the result.
// Use Cases: Enforce corporate branding colors across all charts in automated financial reports. | Standardize visual style of dashboards generated with Aspose.Cells before distribution. | Maintain consistent chart aesthetics when programmatically updating Excel workbooks.
// AI Prompts: Write C# code that changes the second chart's series colors to a different Aspose.Cells palette. | Show how to loop through every chart in a workbook and assign a specific color palette to each series collection. | Explain how to create a custom color palette and apply it to chart series using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Load an XLSX workbook, locate the first chart, apply the MonochromaticPalette6 palette to its series collection, and save the updated file with Aspose.Cells C# API.
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

        Console.WriteLine("Chart series color changed and workbook saved.");
    }
}
