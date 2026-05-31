using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class ChangeChartThemeColors
{
    static void Main()
    {
        // Load the existing XLS workbook
        Workbook workbook = new Workbook("input.xls");

        // Assume the chart is on the first worksheet; adjust index if needed
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure there is at least one chart
        if (worksheet.Charts.Count == 0)
        {
            Console.WriteLine("No charts found in the worksheet.");
            return;
        }

        // Get the first chart (or specify the required chart index)
        Chart chart = worksheet.Charts[0];

        // Access the series collection of the chart
        SeriesCollection seriesColl = chart.NSeries;

        // Change the chart's color palette to MonochromaticPalette6 (Accent6 gradient)
        seriesColl.ChangeColors(ChartColorPaletteType.MonochromaticPalette6);

        // Save the modified workbook
        workbook.Save("output.xls");
        Console.WriteLine("Chart theme colors updated and workbook saved as output.xls.");
    }
}