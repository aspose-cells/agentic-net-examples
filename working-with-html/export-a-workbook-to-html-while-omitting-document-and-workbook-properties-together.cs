// Title: How to export an Excel workbook to HTML while suppressing document and workbook properties using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an XLSX workbook and saves it as HTML, configuring HtmlSaveOptions to omit both document and workbook properties. | Show how to use Aspose.Cells HtmlSaveOptions in C# to generate an HTML file from a workbook while suppressing all metadata.
// Common Searches: Aspose.Cells C# export workbook to HTML without document properties | How to disable ExportWorkbookProperties in HtmlSaveOptions Aspose.Cells | Save Excel as HTML without embedding workbook metadata using Aspose.Cells | C# convert XLSX to HTML and hide workbook properties Aspose.Cells | HtmlSaveOptions ExportDocumentProperties false example
// Tags: Aspose.Cells HtmlSaveOptions disable document metadata | Aspose.Cells HtmlSaveOptions disable workbook metadata | C# Aspose.Cells HTML export without embedded properties | Suppress workbook properties in HTML conversion Aspose.Cells | Convert XLSX to HTML without metadata using Aspose.Cells

using System;
using Aspose.Cells;

// Loads 'input.xlsx', sets HtmlSaveOptions to omit document and workbook properties, and saves as 'output.html' without embedding any workbook or document metadata.
class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to omit document and workbook properties
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
        saveOptions.ExportDocumentProperties = false; // Omit document properties
        saveOptions.ExportWorkbookProperties = false; // Omit workbook properties

        // Save the workbook as HTML with the specified options
        workbook.Save("output.html", saveOptions);
    }
}
