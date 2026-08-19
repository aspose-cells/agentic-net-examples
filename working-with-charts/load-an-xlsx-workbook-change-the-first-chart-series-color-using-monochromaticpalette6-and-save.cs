// Title: C# Example: Apply MonochromaticPalette6 to the First Chart in an XLSX Workbook using Aspose.Cells
// Description: Loads an XLSX file, verifies the first worksheet contains a chart, changes the chart's series colors to the MonochromaticPalette6 palette, and saves the workbook as a new file.
// Keywords: Aspose.Cells C# chart color palette | MonochromaticPalette6 Aspose.Cells | change Excel chart series color .NET | apply monochrome palette to chart | Aspose.Cells chart styling example | C# Excel chart color modification | GitHub Aspose.Cells sample | .NET chart color customization
// Common Searches: Aspose.Cells set MonochromaticPalette6 for chart series | C# change first chart colors in XLSX with Aspose.Cells | how to apply monochrome palette to Excel chart using .NET | example code for chart color palette Aspose.Cells | modify Excel chart series color programmatically
// Developer Intent: Load an XLSX workbook, apply the MonochromaticPalette6 palette to the first chart's series, and save the updated file.
// Use Cases: Create uniform, print‑friendly chart colors for financial reporting dashboards. | Retrofit legacy Excel reports with a consistent monochrome style for corporate branding. | Automate chart styling in batch processes before sharing workbooks with stakeholders.
// AI Prompts: Generate C# code that applies MonochromaticPalette6 to every chart in a workbook using Aspose.Cells. | Add comprehensive error handling for missing worksheets or charts when changing series colors. | Show how to switch a chart back to its original color palette after applying MonochromaticPalette6.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an XLSX file, verifies the first worksheet contains a chart, changes the chart's series colors to the MonochromaticPalette6 palette, and saves the workbook as a new file.
class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Proceed only if the worksheet contains at least one chart
        if (worksheet.Charts.Count > 0)
        {
            // Get the first chart in the worksheet
            Chart chart = worksheet.Charts[0];

            // Retrieve the series collection of the chart
            SeriesCollection seriesCollection = chart.NSeries;

            // Apply the MonochromaticPalette6 color palette to the series
            seriesCollection.ChangeColors(ChartColorPaletteType.MonochromaticPalette6);
        }

        // Save the modified workbook
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
