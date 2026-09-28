// Title: Export an Excel workbook to HTML without workbook properties using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, sets HtmlSaveOptions.ExportWorkbookProperties to false, and saves it as an .html file with Aspose.Cells. | Show how to configure Aspose.Cells HtmlSaveOptions to omit workbook metadata when converting a workbook to HTML in a .NET application.
// Common Searches: Aspose.Cells how to save Excel as HTML without workbook properties | C# HtmlSaveOptions ExportWorkbookProperties false example | convert XLSX to HTML without metadata using Aspose.Cells .NET | disable workbook properties in HTML export Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExportWorkbookProperties | export workbook to HTML without metadata | C# Aspose.Cells HTML conversion omit properties | disable workbook properties Aspose.Cells

using Aspose.Cells;
using System;

// Loads an Excel file, disables ExportWorkbookProperties in HtmlSaveOptions, and saves the workbook as an HTML file, preventing workbook properties from being included in the output.
class Program
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to omit workbook properties
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportWorkbookProperties = false; // Disable exporting of workbook properties

        // Export the workbook to HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
