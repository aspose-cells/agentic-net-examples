using System;
using Aspose.Cells;

namespace ExportWorkbookToHtml
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello World!");

            // Set some workbook properties (optional, they will be omitted in HTML)
            workbook.BuiltInDocumentProperties.Author = "John Doe";
            workbook.BuiltInDocumentProperties.Title = "Sample Workbook";

            // Create HTML save options
            HtmlSaveOptions options = new HtmlSaveOptions();

            // Disable exporting of workbook properties
            options.ExportWorkbookProperties = false;

            // Save the workbook as HTML
            workbook.Save("output.html", options);

            Console.WriteLine("Workbook exported to HTML without workbook properties.");
        }
    }
}