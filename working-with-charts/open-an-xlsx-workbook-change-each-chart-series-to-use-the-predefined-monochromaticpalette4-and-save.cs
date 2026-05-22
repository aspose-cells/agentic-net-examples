using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Loop through each worksheet in the workbook
        foreach (Worksheet ws in workbook.Worksheets)
        {
            // Loop through each chart on the worksheet
            foreach (Chart chart in ws.Charts)
            {
                // Get the series collection of the chart
                SeriesCollection seriesColl = chart.NSeries;

                // Change the color palette of all series to MonochromaticPalette4
                seriesColl.ChangeColors(ChartColorPaletteType.MonochromaticPalette4);
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}