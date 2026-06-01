using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Load the existing XLS workbook
        string inputPath = "input.xls";
        Workbook workbook = new Workbook(inputPath);

        // Choose a monochromatic palette to apply to all chart series
        ChartColorPaletteType palette = ChartColorPaletteType.MonochromaticPalette5;

        // Iterate through each worksheet and each chart within it
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            foreach (Chart chart in sheet.Charts)
            {
                // Apply the selected palette to the chart's series collection
                SeriesCollection seriesColl = chart.NSeries;
                seriesColl.ChangeColors(palette);
            }
        }

        // Save the modified workbook
        string outputPath = "output.xls";
        workbook.Save(outputPath);
    }
}