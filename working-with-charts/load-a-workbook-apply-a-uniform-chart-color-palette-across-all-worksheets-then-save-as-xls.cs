// Title: Set a Uniform Chart Color Palette for All Worksheets and Save as XLS with Aspose.Cells for .NET
// Description: This example loads a workbook, walks through every worksheet and each chart, and applies the same monochromatic ChartColorPaletteType to all series using SeriesCollection.ChangeColors. The workbook is then saved as an Excel 97‑2003 file (XLS) with XlsSaveOptions.MatchColor to map colors to the limited XLS palette.
// Keywords: Aspose.Cells | C# chart color palette | SeriesCollection.ChangeColors | ChartColorPaletteType | XlsSaveOptions | MatchColor | uniform chart colors | Excel 97-2003 export | batch chart formatting | workbook chart styling
// Common Searches: how to apply the same color palette to all charts in an Excel file using Aspose.Cells | Aspose.Cells change chart colors across worksheets | save workbook as XLS while preserving chart colors | C# set chart palette for multiple sheets Aspose.Cells | uniform chart styling for legacy Excel format
// Developer Intent: Apply a single chart color palette to every chart in a workbook and export the result as an XLS file.
// Use Cases: Standardize chart appearance for corporate reports before distributing them in the legacy XLS format. | Prepare workbooks for older systems that only support Excel 97‑2003 by enforcing a consistent monochrome chart style. | Automate batch processing of multiple workbooks to ensure all charts share the same palette, simplifying visual analysis.
// AI Prompts: Write C# code with Aspose.Cells that sets a specific ChartColorPaletteType for all charts in a workbook and saves it as XLS with MatchColor enabled. | Explain the different ChartColorPaletteType enum values and show how to switch from the default palette to another predefined palette. | Provide a step‑by‑step modification to replace the default palette with custom RGB colors for each chart series.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartPaletteDemo
{
    // This example loads a workbook, walks through every worksheet and each chart, and applies the same monochromatic ChartColorPaletteType to all series using SeriesCollection.ChangeColors. The workbook is then saved as an Excel 97‑2003 file (XLS) with XlsSaveOptions.MatchColor to map colors to the limited XLS palette.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook (can be any supported format)
            string sourcePath = "input.xlsx";

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all charts on the current worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    // Get the series collection of the chart
                    SeriesCollection series = chart.NSeries;

                    // Apply a uniform monochromatic color palette to all series.
                    // Using a numeric cast to a valid enum value (0) as demonstrated in the documentation.
                    series.ChangeColors((ChartColorPaletteType)0);
                }
            }

            // Prepare save options for XLS format.
            // MatchColor ensures that colors are mapped to the limited 56‑color palette of XLS.
            XlsSaveOptions saveOptions = new XlsSaveOptions
            {
                MatchColor = true
            };

            // Save the modified workbook as an Excel 97‑2003 file.
            string outputPath = "output.xls";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook saved with uniform chart palette to '{outputPath}'.");
        }
    }
}
