// Title: Export an Excel workbook to HTML while omitting both document and workbook properties using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, configures HtmlSaveOptions to set ExportDocumentProperties and ExportWorkbookProperties to false, and saves the result as an .html file with Aspose.Cells. | Show how to disable metadata export when converting a workbook to HTML with Aspose.Cells, including the necessary property settings in HtmlSaveOptions.
// Common Searches: Aspose.Cells how to hide document properties in HTML export | C# convert Excel to HTML without workbook metadata using Aspose | Set HtmlSaveOptions ExportDocumentProperties false Aspose.Cells example | Remove all properties from HTML output when saving workbook with Aspose.Cells .NET
// Tags: Aspose.Cells HtmlSaveOptions omit document properties | export workbook to html without metadata Aspose.Cells | C# HtmlSaveOptions ExportWorkbookProperties false | convert xlsx to html without workbook properties Aspose

using Aspose.Cells;

// Loads an Excel workbook, configures HtmlSaveOptions to disable ExportDocumentProperties and ExportWorkbookProperties, and saves the workbook as an HTML file without any document or workbook metadata.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Set HTML save options to omit both document and workbook properties
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportDocumentProperties = false; // Do not export document properties
        htmlOptions.ExportWorkbookProperties = false; // Do not export workbook properties

        // Export the workbook to HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
