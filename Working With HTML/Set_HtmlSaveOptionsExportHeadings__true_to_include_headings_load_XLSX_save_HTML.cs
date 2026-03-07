using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options and enable exporting row/column headings
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportHeadings = true; // Include headings in the HTML output

        // Save the workbook as an HTML file using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}