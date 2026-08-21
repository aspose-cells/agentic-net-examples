// Title: Create a PDF/A‑1a compliant PDF from an Excel workbook using Aspose.Cells for .NET (C#)
// Description: This C# example shows how to build a Workbook, add data, configure PdfSaveOptions with PdfCompliance.PdfA1a, and save the file as a PDF/A‑1a document, ensuring archival‑grade compliance.
// Keywords: Aspose.Cells | PDF/A-1a | PdfSaveOptions | PdfCompliance | C# | .NET | archival PDF | Excel to PDF/A | generate PDF/A-1a | PDF/A-1a compliance
// Common Searches: Aspose.Cells generate PDF/A-1a | C# save workbook as PDF/A-1a | PdfSaveOptions compliance PDF/A-1a example | How to create archival PDF from Excel using Aspose | Set PDF/A-1a compliance in Aspose.Cells .NET
// Developer Intent: Set PdfSaveOptions.Compliance to PdfA1a so the exported PDF meets PDF/A‑1a archival standards.
// Use Cases: Produce legally compliant, long‑term storage PDFs from financial spreadsheets. | Export regulatory reports as PDF/A‑1a for document management systems. | Generate PDF/A‑1a invoices that satisfy industry record‑keeping requirements.
// AI Prompts: Write C# code that validates a PDF created with Aspose.Cells against PDF/A‑1a using Aspose.PDF. | Show how to catch and log errors when Aspose.Cells fails to save a workbook as PDF/A‑1a. | Explain how to programmatically confirm PDF/A‑1a compliance after saving a workbook with PdfSaveOptions.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// This C# example shows how to build a Workbook, add data, configure PdfSaveOptions with PdfCompliance.PdfA1a, and save the file as a PDF/A‑1a document, ensuring archival‑grade compliance.
class PdfA1aComplianceDemo
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("PDF/A-1a compliance test");

        // Create PDF save options and set compliance to PDF/A-1a
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.Compliance = PdfCompliance.PdfA1a;

        // Save the workbook as a PDF with the specified compliance level
        workbook.Save("output_PdfA1a.pdf", pdfOptions);
    }
}
