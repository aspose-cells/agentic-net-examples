using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsTitleValidation
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle create rule)
                Workbook workbook = new Workbook();

                // Example: set some data
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sample Data");

                // OPTIONAL: Uncomment to set a title manually
                // workbook.BuiltInDocumentProperties.Title = "My Document Title";

                // Ensure the Title property is not empty before exporting
                string title = workbook.BuiltInDocumentProperties.Title;
                if (string.IsNullOrWhiteSpace(title))
                {
                    // Set a default title if none was provided
                    title = "Untitled Document";
                    workbook.BuiltInDocumentProperties.Title = title;
                }

                // Export to PDF (lifecycle save rule)
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Ensure the document title appears in the PDF window title bar
                    DisplayDocTitle = true
                };
                workbook.Save("ExportedDocument.pdf", pdfOptions);

                // Export to HTML (another example of export)
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Use the workbook title as the HTML page title
                    PageTitle = title,
                    ExportWorkbookProperties = true
                };
                workbook.Save("ExportedDocument.html", htmlOptions);
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.Error.WriteLine($"File not found: {fnfEx.FileName}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}