// Title: Export Excel to PDF with Custom Document Properties using AspNet.Cells (C#)
// Description: Demonstrates how to create a workbook, add string, date, integer and boolean custom document properties, configure PdfSaveOptions.CustomPropertiesExport = Standard, and save the file as a PDF that carries those properties in the PDF Info dictionary.
// Keywords: Aspose.Cells PDF export | custom document properties | PdfSaveOptions | C# Excel to PDF | PDF metadata | Standard custom properties export | Add DateTime property Aspose.Cells | .NET workbook to PDF | Excel metadata in PDF | automated PDF generation
// Common Searches: Aspose.Cells export workbook to PDF with custom metadata | C# add custom document properties to PDF using Aspose.Cells | PdfSaveOptions.CustomPropertiesExport Standard example | How to embed Excel custom properties in PDF Info dictionary | Save Excel as PDF with author and revision data
// Developer Intent: Convert an Excel workbook to PDF while embedding custom document properties as PDF metadata.
// Use Cases: Produce audit‑ready PDFs that include project name, author, and generation timestamp for indexing systems. | Create version‑controlled reports with revision numbers and approval flags to meet compliance requirements. | Automate document pipelines where a DateTime property records the exact moment of PDF generation.
// AI Prompts: Generate C# code that uses Aspose.Cells to export a workbook to PDF and embed custom properties like Project, Author, GeneratedOn, Revision, and Approved. | Explain the effect of setting PdfSaveOptions.CustomPropertiesExport to Standard on the PDF Info dictionary. | Show how to add a Boolean and a DateTime custom document property to a workbook and have them appear as PDF metadata.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfExportWithCustomProperties
{
    // Demonstrates how to create a workbook, add string, date, integer and boolean custom document properties, configure PdfSaveOptions.CustomPropertiesExport = Standard, and save the file as a PDF that carries those properties in the PDF Info dictionary.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the Workbook() constructor rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet and put some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Aspose.Cells PDF Export Demo");
            sheet.Cells["A2"].PutValue("Custom document properties will be included.");

            // Add custom document properties (uses CustomDocumentPropertyCollection.Add rules)
            workbook.CustomDocumentProperties.Add("Project", "PDF Export");
            workbook.CustomDocumentProperties.Add("Author", "John Doe");
            workbook.CustomDocumentProperties.Add("GeneratedOn", DateTime.Now);
            workbook.CustomDocumentProperties.Add("Revision", 3);
            workbook.CustomDocumentProperties.Add("Approved", true);

            // Create PDF save options (uses PdfSaveOptions constructor rule)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Export custom properties as standard entries in the PDF Info dictionary
            pdfOptions.CustomPropertiesExport = PdfCustomPropertiesExport.Standard;

            // Save the workbook to PDF with the specified options (uses Workbook.Save(string, SaveOptions) rule)
            workbook.Save("ExportedWithCustomProperties.pdf", pdfOptions);
        }
    }
}
