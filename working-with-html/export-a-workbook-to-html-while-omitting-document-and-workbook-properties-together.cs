// Title: Export an Excel workbook to HTML while omitting both document and workbook properties with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel to HTML without document properties | How to hide workbook metadata when saving as HTML using Aspose.Cells | C# HtmlSaveOptions ExportDocumentProperties false example | Save Excel as HTML without workbook properties Aspose.Cells .NET | Generate clean HTML from Excel file excluding metadata with Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExportDocumentProperties false | Aspose.Cells ExportWorkbookProperties false HTML conversion | C# Excel to HTML without metadata | HTML export omitting workbook properties Aspose.Cells | Save workbook as HTML without document properties .NET

using Aspose.Cells;

// The example loads an Excel workbook, configures HtmlSaveOptions to set ExportDocumentProperties and ExportWorkbookProperties to false, and saves the file as HTML, resulting in an HTML output that excludes both document and workbook metadata.
class Program
{
    static void Main()
    {
        // Load the source workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Set HTML save options to exclude document and workbook properties
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        htmlOptions.ExportDocumentProperties = false;   // Omit document properties
        htmlOptions.ExportWorkbookProperties = false;   // Omit workbook properties

        // Export the workbook to HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
