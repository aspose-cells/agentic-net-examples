using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to include row and column headings
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportRowColumnHeadings = true; // Preserve worksheet headings

        // Export the workbook to HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}