// Title: C# – Apply MonochromaticPalette4 to All Chart Series in an XLSX Workbook with Aspose.Cells
// Description: Loads an XLSX file, iterates through each worksheet and chart, and uses SeriesCollection.ChangeColors to assign the predefined MonochromaticPalette4 palette to every series before saving the modified workbook.
// Keywords: Aspose.Cells C# chart colors | MonochromaticPalette4 | SeriesCollection.ChangeColors | Excel chart palette programmatically | apply color palette to all charts | chart styling Aspose.Cells | set chart series colors .NET | bulk chart color update
// Common Searches: Aspose.Cells set chart palette C# | change all Excel chart colors to monochrome | apply MonochromaticPalette4 to workbook charts | SeriesCollection.ChangeColors example | how to update chart series colors in Aspose.Cells
// Developer Intent: Assign the MonochromaticPalette4 color scheme to every chart series in a workbook using Aspose.Cells for .NET.
// Use Cases: Create a uniform, print‑friendly look for charts in financial dashboards. | Enforce corporate branding by applying a single palette across all generated reports. | Modernize legacy Excel files with a consistent monochrome style without manual editing.
// AI Prompts: Write C# code that opens an XLSX file with Aspose.Cells, applies ChartColorPaletteType.MonochromaticPalette4 to all chart series, and saves the result. | Explain the behavior and limitations of SeriesCollection.ChangeColors when used with different chart types. | Add comprehensive error handling for workbooks that contain no worksheets, no charts, or unsupported chart formats while applying a color palette. | Suggest performance optimizations for processing very large workbooks with hundreds of charts.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an XLSX file, iterates through each worksheet and chart, and uses SeriesCollection.ChangeColors to assign the predefined MonochromaticPalette4 palette to every series before saving the modified workbook.
class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        string inputFile = "input.xlsx";
        Workbook workbook = new Workbook(inputFile);

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through each chart on the worksheet
            foreach (Chart chart in sheet.Charts)
            {
                // Get the series collection of the chart
                SeriesCollection seriesColl = chart.NSeries;

                // Apply the MonochromaticPalette4 to all series in the collection
                seriesColl.ChangeColors(ChartColorPaletteType.MonochromaticPalette4);
            }
        }

        // Save the modified workbook
        string outputFile = "output.xlsx";
        workbook.Save(outputFile);
    }
}
