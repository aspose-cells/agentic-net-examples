using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to export column widths as scalable units (em/percent)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.WidthScalable = true; // Enables scalable column width in the generated HTML

        // Save the workbook as HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}