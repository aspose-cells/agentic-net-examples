// Title: Export Aspose.Cells Workbook to HTML without Document or Worksheet Properties (C#)
// Description: This example creates a workbook, adds sample data, sets built‑in document properties, and saves the file as HTML using HtmlSaveOptions with ExportDocumentProperties and ExportWorksheetProperties disabled, so the resulting HTML contains no document or sheet metadata.
// Keywords: Aspose.Cells HTML export C# | HtmlSaveOptions ExportDocumentProperties false | ExportWorksheetProperties false | save workbook as HTML without metadata | .NET Excel to HTML conversion | remove Excel properties in HTML output
// Common Searches: Aspose.Cells export to HTML without document properties | C# hide worksheet properties when saving Excel as HTML | How to omit metadata in Aspose.Cells HTML output | HtmlSaveOptions disable property export | Generate clean HTML from Excel using Aspose.Cells
// Developer Intent: Generate an HTML file from a workbook while suppressing all document and worksheet property information.
// Use Cases: Display Excel data on a public website without revealing author or sheet details. | Comply with privacy regulations by stripping metadata from exported HTML reports. | Create lightweight HTML snippets for embedding in web applications.
// AI Prompts: Provide C# code that saves an Aspose.Cells workbook as HTML with both ExportDocumentProperties and ExportWorksheetProperties set to false. | Explain the impact of HtmlSaveOptions settings on the HTML output and how to confirm that no properties are included. | Walk me through the steps to export a workbook to HTML while removing all metadata using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // This example creates a workbook, adds sample data, sets built‑in document properties, and saves the file as HTML using HtmlSaveOptions with ExportDocumentProperties and ExportWorksheetProperties disabled, so the resulting HTML contains no document or sheet metadata.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");

            // Set some document properties (optional, they will be omitted in the HTML)
            workbook.BuiltInDocumentProperties.Author = "John Doe";
            workbook.BuiltInDocumentProperties.Title = "Sample Workbook";

            // Configure HTML save options to exclude document and worksheet properties
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportDocumentProperties = false,   // Omit document properties
                ExportWorksheetProperties = false   // Omit worksheet properties
            };

            // Save the workbook as an HTML file using the configured options
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Workbook exported to HTML without document and worksheet properties.");
        }
    }
}
