// Title: C# – Convert an Aspose.Cells Workbook to PDF and Set the Subject Property
// Description: Creates a new Workbook, optionally adds sample data, assigns the built‑in Subject property (e.g., "FinancialReport2026"), configures PdfSaveOptions to export custom properties as standard entries, and saves the file as a PDF.
// Keywords: Aspose.Cells | C# PDF conversion | Workbook to PDF | Subject property | BuiltInDocumentProperties | PdfSaveOptions | custom properties export | Excel metadata | document categorization | Aspose.Cells example
// Common Searches: Aspose.Cells set Subject property before PDF export | C# convert Excel to PDF with metadata using Aspose.Cells | PdfSaveOptions custom properties export Aspose.Cells | How to add document properties to PDF with Aspose.Cells | Aspose.Cells PDF conversion example C#
// Developer Intent: Set the workbook’s Subject metadata and generate a PDF using Aspose.Cells.
// Use Cases: Produce financial or regulatory PDFs that are automatically indexed by subject in a DMS. | Export Excel reports to PDF while preserving custom properties for compliance audits. | Batch‑process multiple workbooks, assigning a consistent Subject value for archival and searchability.
// AI Prompts: Generate C# code that sets Title, Author, and Subject built‑in properties before saving a workbook as PDF with Aspose.Cells. | Show how to configure PdfSaveOptions to include custom properties and enforce PDF/A compliance during Excel‑to‑PDF conversion. | Explain a script that scans a folder of .xlsx files, derives a Subject from each filename, and saves each as a PDF with the appropriate metadata using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Creates a new Workbook, optionally adds sample data, assigns the built‑in Subject property (e.g., "FinancialReport2026"), configures PdfSaveOptions to export custom properties as standard entries, and saves the file as a PDF.
class ConvertWorkbookToPdfWithSubject
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data (optional)
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample data for PDF conversion");

        // Define the subject property for categorization
        workbook.BuiltInDocumentProperties.Subject = "FinancialReport2026";

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Export custom properties as standard entries (optional but demonstrates usage)
        pdfOptions.CustomPropertiesExport = PdfCustomPropertiesExport.Standard;

        // Save the workbook as a PDF file using the defined options
        workbook.Save("FinancialReport2026.pdf", pdfOptions);
    }
}
