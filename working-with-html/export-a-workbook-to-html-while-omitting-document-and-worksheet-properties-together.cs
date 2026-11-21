// Title: Export an Excel workbook to HTML without document or worksheet properties using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel to HTML without document properties | How to hide worksheet properties in HTML output with Aspose.Cells | Save workbook as HTML excluding metadata using HtmlSaveOptions | C# generate HTML from XLSX without exporting properties Aspose | Remove Excel file properties when converting to HTML with Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExportDocumentProperties false | Aspose.Cells ExportWorksheetProperties false | Excel to HTML conversion without metadata | C# Aspose.Cells HTML export configuration | omit document and worksheet properties in HTML output

using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an Excel file, configures HtmlSaveOptions to set ExportDocumentProperties and ExportWorksheetProperties to false, and saves the workbook as an HTML file, resulting in HTML output that excludes both document and worksheet metadata.
class Program
{
    static void Main()
    {
        // Load the source workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Set HTML save options to exclude document and worksheet properties
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        htmlOptions.ExportDocumentProperties = false;   // Omit document properties
        htmlOptions.ExportWorksheetProperties = false; // Omit worksheet properties

        // Export the workbook to HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
