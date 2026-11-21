// Title: Export an Excel workbook to HTML while suppressing workbook properties with Aspose.Cells in C#
// Common Searches: Aspose.Cells C# export Excel to HTML without workbook metadata | disable workbook properties when saving Excel as HTML using Aspose.Cells | HtmlSaveOptions ExportWorkbookProperties false example code | how to omit workbook properties in HTML output with Aspose.Cells .NET | save workbook as HTML without properties Aspose.Cells tutorial
// Tags: Aspose.Cells HtmlSaveOptions ExportWorkbookProperties | C# export Excel to HTML without metadata | disable workbook properties Aspose.Cells HTML export | Aspose.Cells HTML save options usage | omit workbook properties from HTML output

using Aspose.Cells;

// The sample loads an Excel file, sets HtmlSaveOptions.ExportWorkbookProperties to false, and saves the workbook as an HTML file, ensuring that workbook properties are not included in the generated HTML.
class Program
{
    static void Main()
    {
        // Load the source workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        // Omit workbook properties from the exported HTML
        htmlOptions.ExportWorkbookProperties = false;

        // Export the workbook to HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
