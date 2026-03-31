using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options
        HtmlSaveOptions options = new HtmlSaveOptions();

        // Enable embedding of metadata in the HTML output
        options.ExportWorkbookProperties = true;          // embed workbook properties
        options.ExportWorksheetProperties = true;        // embed worksheet properties
        options.ExportDocumentProperties = true;         // embed document properties
        options.ExportFrameScriptsAndProperties = true; // embed frame scripts and properties

        // Save the workbook as an HTML file with the configured options
        workbook.Save("output.html", options);
    }
}