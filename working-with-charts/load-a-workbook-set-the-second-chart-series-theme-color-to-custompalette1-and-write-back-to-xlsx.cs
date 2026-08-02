using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure the worksheet contains at least one chart
        if (worksheet.Charts.Count > 0)
        {
            // Get the first chart in the worksheet
            Chart chart = worksheet.Charts[0];

            // Get the series collection of the chart
            SeriesCollection seriesColl = chart.NSeries;

            // Apply the first monochromatic palette (CustomPalette1) to the series.
            // This sets the theme color for the series using the Accent1 gradient.
            seriesColl.ChangeColors(ChartColorPaletteType.MonochromaticPalette1);
        }

        // Save the modified workbook back to XLSX format
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}