// Title: Convert an Excel workbook to HTML while suppressing both workbook and worksheet properties using Aspose.Cells for .NET
// AI Prompts: Write C# code that saves an Excel file as HTML with Aspose.Cells, turning off the workbook‑properties and worksheet‑properties export flags. | Demonstrate how to set up HtmlSaveOptions in Aspose.Cells to generate HTML output that contains no workbook or worksheet metadata.
// Common Searches: Aspose.Cells C# export Excel to HTML without any workbook information | How to hide worksheet details in HTML output using Aspose.Cells | HtmlSaveOptions example to omit workbook and worksheet data in .NET | Convert XLSX to HTML while removing all properties with Aspose.Cells | C# Aspose.Cells HTML conversion without workbook or sheet details
// Tags: Aspose.Cells HtmlSaveOptions disable workbook metadata | Aspose.Cells HtmlSaveOptions disable worksheet metadata | C# export Excel to HTML without metadata | Aspose.Cells clean HTML export options | Aspose.Cells omit sheet details in HTML

using Aspose.Cells;

// Loads input.xlsx, configures HtmlSaveOptions to disable workbook and worksheet metadata export, and saves the workbook as output.html.
class Program
{
    static void Main()
    {
        // Load the source workbook
        var workbook = new Workbook("input.xlsx");

        // Configure HTML save options to omit both workbook and worksheet properties
        var htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        htmlOptions.ExportWorkbookProperties = false;   // Do not export workbook properties
        htmlOptions.ExportWorksheetProperties = false; // Do not export worksheet properties

        // Export the workbook to HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
