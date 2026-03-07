using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to hide overlaid content using CrossHideRight
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            HtmlCrossStringType = HtmlCrossType.CrossHideRight
        };

        // Save the workbook as an HTML file with the specified options
        workbook.Save("output.html", saveOptions);
    }
}