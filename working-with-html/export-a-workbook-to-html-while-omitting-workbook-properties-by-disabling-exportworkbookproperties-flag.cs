// Title: Export Aspose.Cells Workbook to HTML without Document Properties (ExportWorkbookProperties = false)
// Description: Creates a workbook, adds sample data, sets author and title, disables HtmlSaveOptions.ExportWorkbookProperties, and saves the file as HTML so the output contains no workbook metadata.
// Keywords: Aspose.Cells | C# | HtmlSaveOptions | ExportWorkbookProperties | HTML export | remove workbook metadata | Excel to HTML | privacy | clean HTML report
// Common Searches: Aspose.Cells export to HTML without properties | HtmlSaveOptions ExportWorkbookProperties false example | C# save Excel as HTML without metadata | How to hide workbook properties in HTML output | Aspose.Cells HTML export privacy
// Developer Intent: Generate an HTML file from an Excel workbook while omitting all built‑in document properties.
// Use Cases: Publish spreadsheet data on a public website without exposing author or title information. | Create lightweight HTML reports for internal dashboards where metadata is unnecessary. | Distribute Excel content as HTML while complying with privacy or data‑protection policies.
// AI Prompts: Provide a C# snippet that saves an Aspose.Cells workbook to HTML with ExportWorkbookProperties set to false. | Explain which HTML elements are removed when HtmlSaveOptions.ExportWorkbookProperties is disabled. | Show how to configure Aspose.Cells to produce clean HTML output that excludes workbook metadata.

using System;
using Aspose.Cells;

namespace ExportWorkbookToHtml
{
    // Creates a workbook, adds sample data, sets author and title, disables HtmlSaveOptions.ExportWorkbookProperties, and saves the file as HTML so the output contains no workbook metadata.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello World!");

            // Set some workbook properties (optional, will be omitted in HTML)
            workbook.BuiltInDocumentProperties.Author = "John Doe";
            workbook.BuiltInDocumentProperties.Title = "Sample Workbook";

            // Create HTML save options and disable exporting workbook properties
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.ExportWorkbookProperties = false;

            // Save the workbook as HTML with the specified options
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Workbook exported to HTML without workbook properties.");
        }
    }
}
