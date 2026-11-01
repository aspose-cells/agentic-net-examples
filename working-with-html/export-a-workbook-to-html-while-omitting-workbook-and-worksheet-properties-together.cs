// Title: Export an Excel workbook to HTML without workbook or worksheet properties using Aspose.Cells for .NET
// AI Prompts: Generate C# code that saves a Workbook as HTML with HtmlSaveOptions set to ExportWorkbookProperties = false and ExportWorksheetProperties = false. | Show how to use Aspose.Cells HtmlSaveOptions to produce HTML output that excludes all workbook‑level and worksheet‑level metadata.
// Common Searches: Aspose.Cells C# export Excel to HTML without workbook metadata | How to disable worksheet properties in HTML output using Aspose.Cells | HtmlSaveOptions ExportWorkbookProperties false C# example | Remove Excel properties when converting to HTML with Aspose.Cells | C# save workbook as HTML with no metadata Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions suppress workbook metadata | Aspose.Cells HtmlSaveOptions suppress worksheet metadata | C# generate HTML from Excel without properties | clean HTML conversion Aspose.Cells | strip workbook and worksheet metadata Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel file, configures HtmlSaveOptions to set ExportWorkbookProperties and ExportWorksheetProperties to false, and saves the workbook as HTML, resulting in an HTML file that contains no workbook‑level or worksheet‑level property information.
class Program
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to omit both workbook and worksheet properties
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        htmlOptions.ExportWorkbookProperties = false;   // Do not export workbook-level properties
        htmlOptions.ExportWorksheetProperties = false; // Do not export worksheet-level properties

        // Export the workbook to HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
