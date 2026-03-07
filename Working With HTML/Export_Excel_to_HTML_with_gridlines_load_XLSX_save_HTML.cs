using System;
using Aspose.Cells;

class ExportExcelToHtmlWithGridlines
{
    static void Main()
    {
        // Load the source XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to include gridlines
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ExportGridLines = true
        };

        // Save the workbook as an HTML file with gridlines
        workbook.Save("output.html", htmlOptions);
    }
}