using System;
using Aspose.Cells;

class HtmlToXlsxConverter
{
    static void Main()
    {
        // Source HTML file path
        string htmlPath = "input.html";

        // Destination XLSX file path
        string xlsxPath = "output.xlsx";

        // Create load options for HTML import
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        // Example option: enable support for <div> tags
        loadOptions.SupportDivTag = true;

        // Load the HTML document into a workbook using the load options
        Workbook workbook = new Workbook(htmlPath, loadOptions);

        // Export the workbook content to an XLSX file
        workbook.Save(xlsxPath);
    }
}