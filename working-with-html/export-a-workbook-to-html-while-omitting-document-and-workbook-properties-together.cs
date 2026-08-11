// Title: Export Aspose.Cells Workbook to HTML without Document or Workbook Properties (C#)
// Description: Learn how to save a workbook as HTML using Aspose.Cells and HtmlSaveOptions with ExportDocumentProperties and ExportWorkbookProperties set to false, producing clean HTML with no embedded metadata.
// Keywords: Aspose.Cells | C# | .NET | HTML export | HtmlSaveOptions | ExportDocumentProperties false | ExportWorkbookProperties false | omit workbook metadata | remove document properties | clean HTML report
// Common Searches: Aspose.Cells export to HTML without metadata | C# HtmlSaveOptions hide document properties | Disable ExportWorkbookProperties in Aspose.Cells | How to remove workbook properties from HTML output | Save Excel as HTML with no author info Aspose
// Developer Intent: Save a workbook as HTML while excluding both document and workbook properties.
// Use Cases: Publish spreadsheet data on public websites without exposing author or title information. | Create lightweight HTML snapshots for automated reporting pipelines. | Batch‑convert Excel files to HTML for archival purposes, ensuring no hidden metadata is retained.
// AI Prompts: Show C# code to export an Aspose.Cells workbook to HTML without document or workbook properties. | Explain how ExportDocumentProperties and ExportWorkbookProperties affect the HTML output in Aspose.Cells. | Provide steps to verify that the generated HTML file contains no embedded workbook metadata.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Learn how to save a workbook as HTML using Aspose.Cells and HtmlSaveOptions with ExportDocumentProperties and ExportWorkbookProperties set to false, producing clean HTML with no embedded metadata.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello World!");

            // Set some document and workbook properties (these will be omitted in the HTML output)
            workbook.BuiltInDocumentProperties.Author = "John Doe";
            workbook.BuiltInDocumentProperties.Title = "Sample Workbook";

            // Configure HTML save options to exclude both document and workbook properties
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportDocumentProperties = false,   // Omit document properties
                ExportWorkbookProperties = false    // Omit workbook properties
            };

            // Save the workbook as an HTML file using the configured options
            workbook.Save("SampleOutput.html", htmlOptions);

            Console.WriteLine("HTML file saved without document and workbook properties.");
        }
    }
}
