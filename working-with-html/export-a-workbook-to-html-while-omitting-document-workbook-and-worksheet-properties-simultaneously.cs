// Title: Export Aspose.Cells Workbook to HTML without Document, Workbook, or Worksheet Properties (C#)
// Description: Creates a workbook, adds sample data, sets built‑in properties, then uses HtmlSaveOptions with ExportDocumentProperties, ExportWorkbookProperties, and ExportWorksheetProperties set to false to save a clean HTML file that contains no metadata.
// Keywords: Aspose.Cells | HtmlSaveOptions | ExportDocumentProperties | ExportWorkbookProperties | ExportWorksheetProperties | C# | .NET | remove metadata | HTML export | Excel to HTML | hide properties
// Common Searches: Aspose.Cells export HTML without properties | How to hide document properties in HTML output using Aspose.Cells | C# save workbook as HTML without workbook metadata | Remove worksheet properties from HTML with Aspose.Cells | Export Excel to clean HTML using Aspose.Cells .NET
// Developer Intent: Generate an HTML representation of an Excel workbook while stripping all document, workbook, and worksheet metadata.
// Use Cases: Publish a privacy‑safe HTML report that does not expose author or title information. | Provide public download links for Excel content where metadata must be omitted for compliance. | Embed lightweight HTML versions of spreadsheets in web applications, reducing page size by excluding property sections.
// AI Prompts: Show a C# example that exports an Aspose.Cells workbook to HTML with ExportDocumentProperties, ExportWorkbookProperties, and ExportWorksheetProperties disabled. | Explain each HtmlSaveOptions flag for removing metadata when converting Excel to HTML with Aspose.Cells. | Suggest additional HtmlSaveOptions to further compress the HTML output after omitting all workbook properties.

using System;
using Aspose.Cells;

namespace ExportHtmlWithoutProperties
{
    // Creates a workbook, adds sample data, sets built‑in properties, then uses HtmlSaveOptions with ExportDocumentProperties, ExportWorkbookProperties, and ExportWorksheetProperties set to false to save a clean HTML file that contains no metadata.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");

            // Set some document and workbook properties (these will be omitted in the HTML output)
            workbook.BuiltInDocumentProperties.Author = "John Doe";
            workbook.BuiltInDocumentProperties.Title = "Demo Workbook";

            // Configure HTML save options to exclude document, workbook, and worksheet properties
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportDocumentProperties = false,   // Omit document properties
                ExportWorkbookProperties = false,   // Omit workbook properties
                ExportWorksheetProperties = false   // Omit worksheet properties
            };

            // Save the workbook as an HTML file using the configured options
            workbook.Save("OutputWithoutProperties.html", htmlOptions);

            Console.WriteLine("HTML file saved without document, workbook, and worksheet properties.");
        }
    }
}
