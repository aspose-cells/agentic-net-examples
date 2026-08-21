// Title: Export Aspose.Cells Workbook to HTML without Workbook or Worksheet Properties (C#)
// Description: Demonstrates saving an Aspose.Cells workbook as HTML while suppressing both workbook‑level and worksheet‑level metadata. Setting HtmlSaveOptions.ExportWorkbookProperties and ExportWorksheetProperties to false produces clean HTML that contains only cell values, perfect for lightweight web reports.
// Keywords: Aspose.Cells HTML export | ExportWorkbookProperties false | ExportWorksheetProperties false | C# Aspose.Cells omit metadata | save workbook as HTML | remove document properties | HTML conversion Aspose.Cells
// Common Searches: Aspose.Cells C# export to HTML without workbook properties | Hide worksheet properties when converting Excel to HTML using Aspose.Cells | C# HtmlSaveOptions exclude document metadata Aspose.Cells | Remove author and sheet name from HTML output Aspose.Cells | Generate clean HTML from Excel with Aspose.Cells C#
// Developer Intent: Create an HTML file from a workbook that excludes all workbook and worksheet metadata.
// Use Cases: Publish Excel data on public websites without revealing author or sheet names. | Produce compact HTML reports for email or intranet portals. | Automate batch conversion of spreadsheets to clean HTML for embedding in web applications.
// AI Prompts: Show C# code that exports an Aspose.Cells workbook to HTML while disabling ExportWorkbookProperties and ExportWorksheetProperties. | Provide an example of using HtmlSaveOptions to omit workbook and worksheet metadata during HTML conversion with Aspose.Cells. | Explain how to configure HtmlSaveOptions in Aspose.Cells to generate HTML without any document properties.

using System;
using Aspose.Cells;

// Demonstrates saving an Aspose.Cells workbook as HTML while suppressing both workbook‑level and worksheet‑level metadata. Setting HtmlSaveOptions.ExportWorkbookProperties and ExportWorksheetProperties to false produces clean HTML that contains only cell values, perfect for lightweight web reports.
class ExportHtmlWithoutProperties
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add some sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello World");
        sheet.Cells["B2"].PutValue(123);

        // (Optional) Set some workbook and worksheet properties to demonstrate they will be omitted
        workbook.BuiltInDocumentProperties.Author = "Demo Author";
        sheet.Name = "DemoSheet";

        // Configure HTML save options to exclude workbook and worksheet properties
        HtmlSaveOptions options = new HtmlSaveOptions();
        options.ExportWorkbookProperties = false;      // Omit workbook properties
        options.ExportWorksheetProperties = false;    // Omit worksheet properties

        // Save the workbook as HTML with the specified options
        workbook.Save("output_without_properties.html", options);
    }
}
