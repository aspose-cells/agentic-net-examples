// Title: Set MonochromaticPalette4 for Every Chart Series in an Excel Workbook with Aspose.Cells (C#)
// Description: Loads an XLSX file, iterates through all worksheets and charts, applies the built‑in MonochromaticPalette4 to each chart's series collection, and saves the updated workbook. Demonstrates Aspose.Cells chart styling in .NET.
// Keywords: Aspose.Cells chart palette | MonochromaticPalette4 C# | change Excel chart colors programmatically | apply color scheme to chart series | Aspose.Cells .NET example | global Excel styling
// Common Searches: how to set a monochrome palette for Excel charts using Aspose.Cells | C# change all chart series colors to MonochromaticPalette4 | iterate worksheets and charts to update series colors Aspose.Cells | apply single color theme to multiple charts in a workbook
// Developer Intent: Programmatically replace the colors of every chart series in a workbook with the MonochromaticPalette4 palette and persist the changes.
// Use Cases: Create a uniform grayscale look for charts in financial reports destined for black‑and‑white printing. | Enforce corporate visual standards by applying a single palette to all charts in a template workbook. | Prepare Excel dashboards for international distribution where a consistent color scheme improves readability.
// AI Prompts: Write C# code using Aspose.Cells that opens an Excel file, loops through all worksheets and charts, sets MonochromaticPalette4 for each chart series, and saves the file. | Show how to add error handling for workbooks that contain no charts while still applying a chosen palette when charts exist. | Explain how to refactor the routine so the palette type is passed as a method argument, allowing any ChartColorPaletteType to be applied.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an XLSX file, iterates through all worksheets and charts, applies the built‑in MonochromaticPalette4 to each chart's series collection, and saves the updated workbook. Demonstrates Aspose.Cells chart styling in .NET.
class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet worksheet in workbook.Worksheets)
        {
            // Iterate through all charts on the current worksheet
            foreach (Chart chart in worksheet.Charts)
            {
                // Get the series collection of the chart
                SeriesCollection seriesCollection = chart.NSeries;

                // Apply the MonochromaticPalette4 to all series in the collection
                seriesCollection.ChangeColors(ChartColorPaletteType.MonochromaticPalette4);
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
