// Title: C# – Apply MonochromaticPalette4 to Every Chart Series in an Excel Workbook with Aspose.Cells
// Description: Load an existing XLSX file using Aspose.Cells for .NET, loop through all worksheets and charts, and use SeriesCollection.ChangeColors with ChartColorPaletteType.MonochromaticPalette4 to recolor every chart series before saving the workbook.
// Keywords: Aspose.Cells | C# Excel chart colors | SeriesCollection.ChangeColors | ChartColorPaletteType | MonochromaticPalette4 | batch update chart palette | .NET Excel chart styling | apply color palette to charts | Excel workbook chart formatting
// Common Searches: How to set MonochromaticPalette4 for all chart series in Excel using Aspose.Cells | C# code to change chart colors to a predefined palette in an XLSX file | Aspose.Cells example for iterating charts and applying a color scheme | Apply a single color palette to every chart in a workbook programmatically | SeriesCollection.ChangeColors usage in Aspose.Cells .NET
// Developer Intent: Recolor every chart series in an existing XLSX workbook to the MonochromaticPalette4 palette and save the updated file.
// Use Cases: Ensure consistent chart appearance across financial reports before distribution. | Prepare presentation‑ready workbooks by applying a uniform monochrome style to all charts. | Automate corporate branding rules by enforcing a specific chart color palette in generated Excel files.
// AI Prompts: Write C# code that opens an XLSX workbook with Aspose.Cells, applies ChartColorPaletteType.MonochromaticPalette4 to all chart series, and saves the file. | Explain the SeriesCollection.ChangeColors method and list all ChartColorPaletteType options available in Aspose.Cells. | Add error handling to the chart‑coloring example for workbooks that contain no charts or have protected sheets.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Load an existing XLSX file using Aspose.Cells for .NET, loop through all worksheets and charts, and use SeriesCollection.ChangeColors with ChartColorPaletteType.MonochromaticPalette4 to recolor every chart series before saving the workbook.
class Program
{
    static void Main()
    {
        // Path to the existing workbook
        string inputPath = "input.xlsx";

        // Load the workbook (lifecycle rule: load)
        Workbook workbook = new Workbook(inputPath);

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all charts on the worksheet
            foreach (Chart chart in sheet.Charts)
            {
                // Get the series collection of the chart
                SeriesCollection seriesColl = chart.NSeries;

                // Apply the MonochromaticPalette4 to all series in the collection
                seriesColl.ChangeColors(ChartColorPaletteType.MonochromaticPalette4);
            }
        }

        // Save the modified workbook (lifecycle rule: save)
        string outputPath = "output.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);

        Console.WriteLine($"Workbook saved to '{outputPath}'.");
    }
}
