using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data
            sheet.Cells["A1"].PutValue("Hello World!");
            sheet.Cells["B2"].PutValue(12345);
            sheet.Cells["C3"].PutValue(DateTime.Now);

            // Set custom document properties (author, title, etc.)
            workbook.BuiltInDocumentProperties.Author = "John Doe";
            workbook.BuiltInDocumentProperties.Title = "Sample HTML Export";
            workbook.BuiltInDocumentProperties.Subject = "Aspose.Cells Demo";

            // Create HTML save options (default constructor)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Ensure workbook properties are exported (default is true, set explicitly)
            htmlOptions.ExportWorkbookProperties = true;

            // Keep other default CSS settings (do not disable CSS, keep unused styles excluded)
            // No need to modify DisableCss or ExcludeUnusedStyles as defaults are suitable

            // Save the workbook as HTML with the specified options
            string outputPath = "ExportedWorkbook.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook exported to HTML at: {outputPath}");
        }
    }
}