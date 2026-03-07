using System;
using Aspose.Cells;

class ExportExcelToHtmlWithBlankLinks
{
    static void Main()
    {
        // Load the Excel workbook from an existing XLSX file
        Workbook workbook = new Workbook("input.xlsx");

        // Initialize HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();

        // Set links to open in a new window or tab (_blank)
        saveOptions.LinkTargetType = HtmlLinkTargetType.Blank;

        // Save the workbook as an HTML file using the configured options
        workbook.Save("output.html", saveOptions);
    }
}