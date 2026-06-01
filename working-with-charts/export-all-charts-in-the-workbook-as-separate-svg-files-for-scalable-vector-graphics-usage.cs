using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class ExportChartsToSvg
{
    static void Main()
    {
        // Load the workbook from a file (replace with your actual path)
        Workbook workbook = new Workbook("input.xlsx");

        // Loop through each worksheet in the workbook
        for (int sheetIndex = 0; sheetIndex < workbook.Worksheets.Count; sheetIndex++)
        {
            Worksheet sheet = workbook.Worksheets[sheetIndex];

            // Loop through each chart on the current worksheet
            for (int chartIndex = 0; chartIndex < sheet.Charts.Count; chartIndex++)
            {
                Chart chart = sheet.Charts[chartIndex];

                // Create a unique file name for each chart
                string svgFileName = $"Chart_Sheet{sheetIndex}_Chart{chartIndex}.svg";

                // Export the chart to SVG using the built‑in overload that accepts ImageType
                chart.ToImage(svgFileName, ImageType.Svg);
            }
        }

        Console.WriteLine("All charts have been exported as separate SVG files.");
    }
}