using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options and enable exporting comments as tooltips
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.IsExportComments = true; // comments will appear as tooltips in the HTML output

        // Save the workbook as an HTML file using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}