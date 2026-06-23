using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure the worksheet contains at least one chart
        if (worksheet.Charts.Count == 0)
        {
            Console.WriteLine("No charts found in the worksheet.");
            return;
        }

        // Get the first chart (replace with the appropriate index if necessary)
        Chart chart = worksheet.Charts[0];

        // Access the series collection of the chart
        SeriesCollection seriesCollection = chart.NSeries;

        // Verify that there are at least three series in the chart
        if (seriesCollection.Count < 3)
        {
            Console.WriteLine("The chart does not contain three series.");
            return;
        }

        // Aspose.Cells does not expose a per‑series palette setter.
        // As a workaround, apply the desired palette to the whole collection.
        // The enum value for CustomPalette2 is not defined in older versions,
        // so we cast the integer value that represents it (e.g., 2).
        ChartColorPaletteType customPalette2 = (ChartColorPaletteType)2; // Represents CustomPalette2

        // Apply the custom palette to the series collection.
        // This will affect all series, including the third one.
        seriesCollection.ChangeColors(customPalette2);

        // Save the modified workbook
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}