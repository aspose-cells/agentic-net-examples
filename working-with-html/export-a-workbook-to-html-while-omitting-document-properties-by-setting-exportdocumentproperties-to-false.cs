// Title: Export Aspose.Cells Workbook to HTML without Document Properties (ExportDocumentProperties = false)
// Description: Shows how to save a workbook as an HTML file using Aspose.Cells for .NET while suppressing built‑in document properties by configuring HtmlSaveOptions.ExportDocumentProperties to false.
// Keywords: Aspose.Cells | HTML export | ExportDocumentProperties | C# | .NET | remove metadata | workbook to HTML | HtmlSaveOptions | document properties | privacy
// Common Searches: Aspose.Cells export HTML without document properties | HtmlSaveOptions ExportDocumentProperties false example | C# save Excel as HTML without metadata | remove author and title from Aspose.Cells HTML output | how to hide workbook properties when converting to HTML
// Developer Intent: Save an Excel workbook as HTML while omitting all built‑in document properties.
// Use Cases: Create public‑facing web pages from spreadsheets without exposing author or title information. | Generate clean HTML reports for intranet portals while complying with data‑privacy policies. | Automate batch conversion of internal Excel files to HTML without leaking metadata.
// AI Prompts: Provide a C# snippet that converts an Aspose.Cells workbook to HTML and disables document property export. | Explain how to configure HtmlSaveOptions in Aspose.Cells to exclude metadata when saving as HTML. | Show the steps to hide author and title information in the HTML output of an Excel workbook using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExportHtmlWithoutDocProps
{
    // Shows how to save a workbook as an HTML file using Aspose.Cells for .NET while suppressing built‑in document properties by configuring HtmlSaveOptions.ExportDocumentProperties to false.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");

            // Set some document properties (these will be omitted in the HTML output)
            workbook.BuiltInDocumentProperties.Author = "John Doe";
            workbook.BuiltInDocumentProperties.Title = "Sample Workbook";

            // Configure HTML save options to exclude document properties
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.ExportDocumentProperties = false; // Omit document properties

            // Save the workbook as an HTML file using the configured options
            workbook.Save("output_without_doc_props.html", htmlOptions);

            Console.WriteLine("HTML file saved without document properties.");
        }
    }
}
