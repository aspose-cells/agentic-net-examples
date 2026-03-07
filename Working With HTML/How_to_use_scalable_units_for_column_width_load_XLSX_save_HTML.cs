using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ScalableColumnWidthDemo
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            // Load an existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Example: get the width of the first column in pixels (original width)
            double widthInPixels = workbook.Worksheets[0].Cells.GetColumnWidth(0, true, CellsUnitType.Pixel);
            Console.WriteLine($"Column A width in pixels before scaling: {widthInPixels}");

            // Save the workbook as HTML with default (scalable) column widths
            HtmlSaveOptions scalableOptions = new HtmlSaveOptions();
            workbook.Save("output_scalable.html", scalableOptions);

            // Save another HTML file with default (fixed) column widths for comparison
            HtmlSaveOptions fixedOptions = new HtmlSaveOptions();
            workbook.Save("output_fixed.html", fixedOptions);
        }
    }
}