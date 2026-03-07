using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to include document, workbook, and worksheet properties
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportDocumentProperties = true;   // Export built‑in document properties
        htmlOptions.ExportWorkbookProperties = true;   // Export workbook‑level properties
        htmlOptions.ExportWorksheetProperties = true; // Export each worksheet's properties

        // Save the workbook as an HTML file with the specified options
        workbook.Save("output.html", htmlOptions);
    }
}