using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

class ApplyUniformChartPalette
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        string inputFile = "input.xlsx";
        Workbook workbook = new Workbook(inputFile);

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all charts in the worksheet
            foreach (Chart chart in sheet.Charts)
            {
                // Apply a uniform monochromatic color palette to the chart series.
                // Using a numeric cast to a valid ChartColorPaletteType value (e.g., 0).
                chart.NSeries.ChangeColors((ChartColorPaletteType)0);
            }
        }

        // Prepare XLS save options (ensure colors are matched to the 56‑color palette)
        XlsSaveOptions saveOptions = new XlsSaveOptions();
        saveOptions.MatchColor = true;

        // Save the workbook as an Excel 97‑2003 file
        workbook.Save("output.xls", saveOptions);
    }
}