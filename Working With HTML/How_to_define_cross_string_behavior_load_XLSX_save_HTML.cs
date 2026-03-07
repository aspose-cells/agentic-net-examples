using System;
using Aspose.Cells;

namespace AsposeCellsCrossStringDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook from file
            Workbook workbook = new Workbook("input.xlsx");

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Define cross‑string behavior.
            // Options:
            //   HtmlCrossType.Default      – Excel‑like behavior
            //   HtmlCrossType.MSExport    – Same as Excel export
            //   HtmlCrossType.Cross       – Faster for large files, always cross cells
            //   HtmlCrossType.CrossHideRight – Hide right‑hand text when overlapping
            //   HtmlCrossType.FitToCell   – Truncate to cell width
            htmlOptions.HtmlCrossStringType = HtmlCrossType.Cross;

            // Save the workbook as HTML using the defined options
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Workbook saved to HTML with cross‑string behavior set to 'Cross'.");
        }
    }
}