// Title: Aspose.Cells C# Export Workbook to HTML without Document or Workbook Properties
// Description: Demonstrates how to save an Aspose.Cells workbook as an HTML file while suppressing both built‑in document properties and custom workbook properties using HtmlSaveOptions.
// Keywords: Aspose.Cells | C# | HTML export | ExportDocumentProperties false | ExportWorkbookProperties false | remove Excel metadata | HtmlSaveOptions | privacy‑focused HTML conversion | Excel to HTML without properties
// Common Searches: Aspose.Cells export to HTML without metadata | C# HtmlSaveOptions hide document properties | disable workbook properties in HTML output Aspose | save Excel as HTML no document properties | Aspose.Cells HTML conversion privacy
// Developer Intent: Generate an HTML representation of a workbook while omitting all document and workbook metadata.
// Use Cases: Create clean HTML reports that do not reveal Excel file metadata. | Publish Excel data on public websites where privacy compliance is required. | Automate bulk conversion of workbooks to HTML with consistent metadata suppression.
// AI Prompts: Show C# code that exports an Aspose.Cells workbook to HTML with ExportDocumentProperties and ExportWorkbookProperties set to false. | Explain the impact of HtmlSaveOptions ExportDocumentProperties and ExportWorkbookProperties on the resulting HTML file. | Provide a script to batch‑convert multiple Excel files to HTML using Aspose.Cells, ensuring no metadata is included.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Demonstrates how to save an Aspose.Cells workbook as an HTML file while suppressing both built‑in document properties and custom workbook properties using HtmlSaveOptions.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello World!");

            // Configure HTML save options to omit both document and workbook properties
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportDocumentProperties = false, // Do not export built‑in document properties
                ExportWorkbookProperties = false   // Do not export workbook properties
            };

            // Save the workbook as an HTML file using the configured options
            workbook.Save("output.html", htmlOptions);
        }
    }
}
