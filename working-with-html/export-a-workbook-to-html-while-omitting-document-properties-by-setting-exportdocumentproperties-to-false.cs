// Title: Saving an Excel workbook as HTML without document properties using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, sets HtmlSaveOptions so that document properties are excluded, and saves the workbook as an .html file using Aspose.Cells. | Show a step‑by‑step example of using Aspose.Cells HtmlSaveOptions to export Excel to HTML while suppressing metadata.
// Common Searches: aspnet save excel as html without document properties aspose.cells | c# export workbook to html exclude metadata aspose | how to disable ExportDocumentProperties in Aspose.Cells HtmlSaveOptions | remove Excel file properties from generated HTML using Aspose.Cells
// Tags: disable document properties in HTML export | C# convert Excel to HTML omitting metadata | Aspose.Cells .NET generate clean HTML from workbook | Export Excel workbook as HTML omitting properties

using Aspose.Cells;
using System;

// // Loads an Excel workbook, configures HtmlSaveOptions to exclude document properties, and saves the workbook as a clean HTML file without metadata.
class Program
{
    static void Main()
    {
        // Load the source workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Set up HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        // Omit document properties from the exported HTML
        htmlOptions.ExportDocumentProperties = false;

        // Export the workbook to HTML
        workbook.Save("output.html", htmlOptions);
    }
}
