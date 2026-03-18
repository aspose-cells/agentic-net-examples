using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Load the source XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Suppress any overlaid elements (shapes, charts, pictures) by ignoring invisible shapes
        // and preventing export of hidden worksheets which may contain such elements.
        htmlOptions.IgnoreInvisibleShapes = true;      // Do not export shapes that are not visible
        htmlOptions.ExportHiddenWorksheet = false;    // Skip hidden worksheets entirely

        // Save the workbook as HTML with the configured options
        workbook.Save("output.html", htmlOptions);
    }
}