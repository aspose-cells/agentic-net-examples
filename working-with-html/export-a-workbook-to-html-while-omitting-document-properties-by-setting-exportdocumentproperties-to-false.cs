// Title: Export an Excel workbook to HTML without document properties using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export workbook to HTML without document properties | How to disable ExportDocumentProperties when saving Excel as HTML in .NET | Save Excel file as HTML with metadata excluded using Aspose.Cells | C# HtmlSaveOptions ExportDocumentProperties false example | Convert Excel to HTML without including workbook properties Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExportDocumentProperties false | C# export Excel to HTML without metadata | disable document properties in Aspose.Cells HTML conversion | Excel to HTML conversion without properties Aspose.Cells | HtmlSaveOptions omit document properties C#

using Aspose.Cells;
using System;

// Demonstrates how to create or load a Workbook, set HtmlSaveOptions.ExportDocumentProperties to false, and save the workbook as an HTML file, thereby omitting all document properties from the generated HTML output.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Populate the workbook with sample data (optional)
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample");
        sheet.Cells["B1"].PutValue("Data");

        // Configure HTML save options to omit document properties
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportDocumentProperties = false; // Do not include document properties in the HTML output

        // Export the workbook to HTML
        workbook.Save("ExportedWorkbook.html", htmlOptions);
    }
}
