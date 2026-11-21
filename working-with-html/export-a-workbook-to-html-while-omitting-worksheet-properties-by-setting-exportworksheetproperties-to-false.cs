// Title: Export an Excel workbook to HTML while omitting worksheet properties with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export workbook to HTML without worksheet properties | HtmlSaveOptions ExportWorksheetProperties false example | How to hide sheet metadata when converting Excel to HTML using Aspose.Cells | Save Excel file as HTML without sheet properties in .NET | Aspose.Cells HTML conversion omit worksheet properties setting
// Tags: Aspose.Cells HtmlSaveOptions configuration | disable worksheet properties in HTML export | C# Excel to HTML conversion settings | omit sheet metadata Aspose.Cells | HTML save options for workbook

using Aspose.Cells;

// Loads an Excel workbook, sets HtmlSaveOptions.ExportWorksheetProperties to false, and saves the file as HTML, resulting in HTML output that does not include worksheet property metadata.
class Program
{
    static void Main()
    {
        // Load the workbook from a file (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        // Omit worksheet properties in the generated HTML
        htmlOptions.ExportWorksheetProperties = false;

        // Export the workbook to HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
