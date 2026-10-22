// Title: Export an Excel workbook to HTML without document or worksheet properties using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an XLSX file and saves it as HTML with HtmlSaveOptions configured to exclude both document and worksheet properties. | Show how to set Aspose.Cells HtmlSaveOptions to suppress all workbook and worksheet metadata during HTML conversion in a .NET application.
// Common Searches: Aspose.Cells C# export Excel to HTML without including document properties | How to prevent worksheet metadata from appearing in HTML output with Aspose.Cells | Example of HtmlSaveOptions to hide workbook properties during HTML conversion | Remove Excel file metadata when generating HTML using Aspose.Cells .NET
// Tags: Aspose.Cells HtmlSaveOptions ExportDocumentProperties false | Aspose.Cells HtmlSaveOptions ExportWorksheetProperties false | C# Excel to HTML conversion without metadata | Suppress workbook properties in HTML export Aspose.Cells | HTML export options for Excel using Aspose.Cells

using Aspose.Cells;

// Loads an XLSX workbook, configures HtmlSaveOptions to omit document and worksheet properties, and saves the workbook as an HTML file.
class Program
{
    static void Main()
    {
        // Load the source workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML export options to omit both document and worksheet properties
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportDocumentProperties = false;   // Do not include document properties
        htmlOptions.ExportWorksheetProperties = false; // Do not include worksheet properties

        // Export the workbook to HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
