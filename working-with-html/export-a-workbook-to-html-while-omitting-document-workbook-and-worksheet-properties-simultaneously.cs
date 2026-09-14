// Title: Export an Excel workbook to HTML without any document, workbook, or worksheet properties using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an XLSX file and saves it as HTML with HtmlSaveOptions configured to disable ExportDocumentProperties, ExportWorkbookProperties, and ExportWorksheetProperties. | Show how to use Aspose.Cells HtmlSaveOptions to produce HTML output that excludes all workbook metadata. | Provide a C# snippet that converts a spreadsheet to HTML while suppressing document, workbook, and worksheet property export.
// Common Searches: Aspose.Cells C# export Excel to HTML without metadata | How to disable document properties in HTML conversion with Aspose.Cells | Save workbook as HTML and hide worksheet properties using Aspose.Cells .NET | HtmlSaveOptions ExportWorkbookProperties false example in C# | Convert XLSX to HTML omitting all properties Aspose.Cells
// Tags: htmlsaveoptions disable metadata export | aspocells export workbook to html without properties | c# convert xlsx to html hide workbook metadata | aspocells omit worksheet properties in html output | aspocells html conversion suppress document metadata

using Aspose.Cells;
using System;

// Loads 'input.xlsx', configures HtmlSaveOptions to turn off ExportDocumentProperties, ExportWorkbookProperties, and ExportWorksheetProperties, and saves the workbook as 'output.html' using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the source workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Set HTML save options to omit all properties
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        htmlOptions.ExportDocumentProperties = false;   // Omit document properties
        htmlOptions.ExportWorkbookProperties = false;   // Omit workbook properties
        htmlOptions.ExportWorksheetProperties = false; // Omit worksheet properties

        // Export the workbook to HTML with the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
