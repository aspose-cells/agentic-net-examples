using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Set HTML save options to hide overlaid (cross) content on the right side
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.HtmlCrossStringType = HtmlCrossType.CrossHideRight;

        // Save the workbook as an HTML file using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}