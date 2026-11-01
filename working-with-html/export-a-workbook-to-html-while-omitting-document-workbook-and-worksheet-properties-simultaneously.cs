// Title: Export an Excel workbook to HTML while suppressing document, workbook, and worksheet properties using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file and saves it as HTML with HtmlSaveOptions configured to turn off ExportDocumentProperties, ExportWorkbookProperties, and ExportWorksheetProperties. | Demonstrate how to use Aspose.Cells HtmlSaveOptions to generate HTML output that excludes all Excel metadata.
// Common Searches: Aspose.Cells how to save Excel as HTML without any properties | C# HtmlSaveOptions ExportDocumentProperties false example | Remove workbook and worksheet metadata when converting to HTML with Aspose.Cells | Export Excel to HTML omitting document properties using .NET SDK | Aspose.Cells HTML conversion hide Excel metadata
// Tags: Aspose.Cells HtmlSaveOptions metadata suppression | C# export Excel to HTML without properties | HtmlSaveOptions ExportDocumentProperties false | Aspose.Cells omit workbook properties in HTML | Convert .xlsx to HTML without worksheet metadata

using Aspose.Cells;
using System;

// // Loads an Excel workbook, configures HtmlSaveOptions to disable exporting of document, workbook, and worksheet properties, and saves the workbook as an HTML file.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Set HTML save options to omit all properties
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        htmlOptions.ExportDocumentProperties = false;   // Omit document properties
        htmlOptions.ExportWorkbookProperties = false;   // Omit workbook properties
        htmlOptions.ExportWorksheetProperties = false; // Omit worksheet properties

        // Export the workbook to HTML
        workbook.Save("output.html", htmlOptions);
    }
}
