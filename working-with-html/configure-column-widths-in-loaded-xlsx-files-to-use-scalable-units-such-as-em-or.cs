using System;
using Aspose.Cells;

namespace AsposeCellsColumnWidthScalableDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Optionally adjust column widths before export (example: set column A to 20 characters)
            // This step is optional and demonstrates that column widths can be customized.
            workbook.Worksheets[0].Cells.Columns[0].Width = 20;

            // Create HTML save options and enable scalable column widths.
            // When WidthScalable is true, column widths are exported using scalable units (e.g., em, %).
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.WidthScalable = true;

            // Save the workbook as HTML with scalable column widths.
            workbook.Save("output_scalable.html", htmlOptions);
        }
    }
}