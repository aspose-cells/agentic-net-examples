using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (index 0)
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure the worksheet contains at least one chart
        if (worksheet.Charts.Count > 0)
        {
            // Get the first chart in the worksheet
            Chart chart = worksheet.Charts[0];

            // Apply the MonochromaticPalette6 to the chart's series collection
            chart.NSeries.ChangeColors(ChartColorPaletteType.MonochromaticPalette6);
        }

        // Save the modified workbook to a new file
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}