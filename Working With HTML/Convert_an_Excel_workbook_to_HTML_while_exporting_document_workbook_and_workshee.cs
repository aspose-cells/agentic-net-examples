using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main()
        {
            // Load the source Excel workbook (XLSX format)
            Workbook workbook = new Workbook("input.xlsx");

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Enable exporting of document, workbook, and worksheet properties
            htmlOptions.ExportDocumentProperties = true;
            htmlOptions.ExportWorkbookProperties = true;
            htmlOptions.ExportWorksheetProperties = true;

            // Save the workbook as HTML with the specified options
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Workbook successfully converted to HTML with properties exported.");
        }
    }
}