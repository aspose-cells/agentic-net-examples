// Title: Save an Excel workbook as HTML without worksheet properties using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, sets HtmlSaveOptions.ExportWorksheetProperties = false, and saves the workbook as .html with Aspose.Cells. | Show how to disable worksheet property export when converting Excel to HTML in a .NET application using Aspose.Cells.
// Common Searches: Aspose.Cells how to export Excel to HTML without worksheet properties | C# HtmlSaveOptions ExportWorksheetProperties false example | Convert .xlsx to .html without sheet metadata using Aspose.Cells | Disable worksheet properties in HTML output Aspose.Cells .NET | Minimal HTML export of Excel workbook Aspose.Cells
// Tags: Aspose.Cells HTML export without worksheet properties | HtmlSaveOptions ExportWorksheetProperties false C# | Excel to HTML conversion minimal metadata Aspose.Cells | C# save workbook as HTML without sheet properties | Aspose.Cells HTMLSaveOptions configuration example

using System;
using Aspose.Cells;

// Loads an Excel file, configures HtmlSaveOptions.ExportWorksheetProperties to false, and saves the workbook as an HTML file, removing worksheet property information from the generated output.
class ExportWorkbookToHtml
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            // Omit worksheet properties from the exported HTML
            ExportWorksheetProperties = false
        };

        // Export the workbook to HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
