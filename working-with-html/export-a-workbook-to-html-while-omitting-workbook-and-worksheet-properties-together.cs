// Title: Export Workbook to HTML without Workbook or Worksheet Properties using Aspose.Cells (C#)
// Description: Shows how to save an Aspose.Cells workbook as HTML while suppressing both workbook and worksheet metadata by setting HtmlSaveOptions.ExportWorkbookProperties and ExportWorksheetProperties to false.
// Keywords: Aspose.Cells | C# | HtmlSaveOptions | ExportWorkbookProperties | ExportWorksheetProperties | HTML export | remove metadata | Excel to HTML | omit workbook properties | omit worksheet properties
// Common Searches: Aspose.Cells export HTML without properties | How to hide workbook metadata in HTML output | C# save Excel as HTML without worksheet properties | HtmlSaveOptions ExportWorkbookProperties false example | Remove document properties from Aspose.Cells HTML export
// Developer Intent: Save a workbook as HTML while excluding all workbook and worksheet properties.
// Use Cases: Public web reports that must not reveal author or title information. | Embedding Excel data in email newsletters without exposing document metadata. | Generating SEO‑friendly HTML pages from Excel data with a clean markup. | Batch conversion of Excel files to HTML for archival or publishing where metadata must be stripped.
// AI Prompts: Generate C# code using Aspose.Cells to export a workbook to HTML with ExportWorkbookProperties and ExportWorksheetProperties disabled. | Explain how HtmlSaveOptions.ExportWorkbookProperties and ExportWorksheetProperties affect the HTML output and how to verify that the metadata is omitted. | Provide a step‑by‑step guide to batch‑convert multiple worksheets to separate HTML files while stripping all workbook and worksheet properties.

using System;
using Aspose.Cells;

// Shows how to save an Aspose.Cells workbook as HTML while suppressing both workbook and worksheet metadata by setting HtmlSaveOptions.ExportWorkbookProperties and ExportWorksheetProperties to false.
class ExportHtmlWithoutProperties
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Hello World!");

        // (Optional) Set workbook properties that we intend to omit in the HTML output
        workbook.BuiltInDocumentProperties.Author = "Sample Author";
        workbook.BuiltInDocumentProperties.Title = "Sample Title";

        // Configure HTML save options to exclude both workbook and worksheet properties
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ExportWorkbookProperties = false,   // Omit workbook properties
            ExportWorksheetProperties = false   // Omit worksheet properties
        };

        // Save the workbook as an HTML file using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
