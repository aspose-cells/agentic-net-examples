// Title: Export an Excel workbook to HTML without workbook or worksheet properties using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export Excel to HTML without workbook properties | How to omit worksheet properties when saving XLSX as HTML with Aspose.Cells | HtmlSaveOptions ExportWorkbookProperties false example in .NET | Save Excel file as HTML without any metadata using Aspose.Cells | C# convert workbook to HTML excluding workbook and worksheet properties
// Tags: Aspose.Cells HtmlSaveOptions ExportWorkbookProperties false | C# HtmlSaveOptions ExportWorksheetProperties false | HTML conversion of Excel without metadata Aspose | Export Excel to HTML omitting workbook properties .NET | Save workbook as HTML with property exclusion Aspose.Cells | Aspose.Cells HTML export options for Excel files

using Aspose.Cells;

// Loads an existing XLSX workbook, configures HtmlSaveOptions to set ExportWorkbookProperties and ExportWorksheetProperties to false, and saves the workbook as an HTML file, thereby removing all workbook and worksheet metadata from the output.
class Program
{
    static void Main()
    {
        // Load an existing workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Set HTML save options to exclude workbook and worksheet properties
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        htmlOptions.ExportWorkbookProperties = false;   // Omit workbook properties
        htmlOptions.ExportWorksheetProperties = false; // Omit worksheet properties

        // Export the workbook to HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
