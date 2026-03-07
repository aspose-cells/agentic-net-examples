using System;
using Aspose.Cells;

class ExportExcelToHtmlWithCssCustomProperties
{
    static void Main()
    {
        // Load the source XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to enable CSS custom properties
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.EnableCssCustomProperties = true;

        // Save the workbook as an HTML file using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}