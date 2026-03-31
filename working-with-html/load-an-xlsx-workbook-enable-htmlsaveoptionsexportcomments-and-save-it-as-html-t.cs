using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options and enable comment export
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            IsExportComments = true
        };

        // Save the workbook as HTML with comments retained
        workbook.Save("output.html", htmlOptions);
    }
}