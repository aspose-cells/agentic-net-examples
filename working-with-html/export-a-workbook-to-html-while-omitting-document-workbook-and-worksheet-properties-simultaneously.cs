// Title: Export an Excel workbook to HTML without document, workbook, or worksheet properties using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export Excel to HTML without document properties | Save workbook as HTML and hide workbook metadata with Aspose.Cells | Aspose.Cells HtmlSaveOptions disable ExportWorksheetProperties example | Remove all properties when converting .xlsx to .html using Aspose.Cells | How to prevent Aspose.Cells from exporting workbook and worksheet properties to HTML | Aspose.Cells C# generate clean HTML from Excel without metadata
// Tags: Aspose.Cells HtmlSaveOptions disable metadata | export Excel to HTML without properties C# | Aspose.Cells ExportDocumentProperties false | HTML conversion omit workbook properties | C# generate HTML from workbook without worksheet metadata

using Aspose.Cells;

// Loads an existing Excel workbook, configures HtmlSaveOptions to set ExportDocumentProperties, ExportWorkbookProperties, and ExportWorksheetProperties to false, and saves the workbook as an HTML file, omitting all document, workbook, and worksheet metadata.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to omit all properties
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportDocumentProperties = false;   // Do not export document properties
        htmlOptions.ExportWorkbookProperties = false;   // Do not export workbook properties
        htmlOptions.ExportWorksheetProperties = false; // Do not export worksheet properties

        // Save the workbook as HTML with the specified options
        workbook.Save("output.html", htmlOptions);
    }
}
