using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

class ReplaceChartSeriesColors
{
    static void Main()
    {
        // Load an existing workbook (lifecycle rule: load)
        Workbook workbook = new Workbook("InputWorkbook.xlsx");

        // Define custom colors to be placed into the palette (optional step)
        // Here we replace palette entries 0, 1, 2 with custom RGB values
        workbook.ChangePalette(Color.FromArgb(255, 255, 128, 0), 0); // Orange
        workbook.ChangePalette(Color.FromArgb(255, 0, 128, 255), 1); // Light Blue
        workbook.ChangePalette(Color.FromArgb(255, 128, 0, 255), 2); // Purple

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all charts in the worksheet
            foreach (Chart chart in sheet.Charts)
            {
                // Get the series collection of the chart
                SeriesCollection seriesColl = chart.NSeries;

                // Apply a monochromatic palette to the series collection
                // This uses the rule: SeriesCollection.ChangeColors
                seriesColl.ChangeColors(ChartColorPaletteType.MonochromaticPalette5);
            }
        }

        // Save the modified workbook (lifecycle rule: save)
        workbook.Save("OutputWorkbook.xlsx", SaveFormat.Xlsx);
    }
}