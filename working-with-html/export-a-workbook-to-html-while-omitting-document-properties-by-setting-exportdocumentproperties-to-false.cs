using System;
using Aspose.Cells;

namespace AsposeCellsExportHtml
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello World!");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Set some document properties (these will be omitted in the HTML output)
            workbook.BuiltInDocumentProperties.Author = "John Doe";
            workbook.BuiltInDocumentProperties.Title = "Sample Workbook";

            // Create HTML save options and disable exporting of document properties
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.ExportDocumentProperties = false; // Omit document properties

            // Save the workbook as an HTML file using the specified options
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Workbook exported to HTML without document properties.");
        }
    }
}