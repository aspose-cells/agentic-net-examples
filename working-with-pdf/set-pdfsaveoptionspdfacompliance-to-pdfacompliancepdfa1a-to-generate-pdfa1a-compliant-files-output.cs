// Title: Generate PDF/A‑1a compliant PDF from an Excel workbook using Aspose.Cells (.NET C#)
// Description: C# code that creates a Workbook, writes a value to cell A1, configures PdfSaveOptions.Compliance = PdfCompliance.PdfA1a, and saves the workbook as a PDF/A‑1a‑compatible file (PdfA1aOutput.pdf) with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | PdfSaveOptions | PDF/A-1a | C# .NET | Excel to PDF/A | PdfCompliance.PdfA1a | archival PDF | document preservation | long‑term PDF compliance
// Common Searches: Aspose.Cells save workbook as PDF/A-1a C# | PdfSaveOptions compliance PDF/A-1a example | C# export Excel to PDF/A-1a | How to create PDF/A-1a with Aspose.Cells | PDF/A-1a conversion .NET
// Developer Intent: Produce a PDF file that meets the PDF/A‑1a archival standard directly from an Excel workbook.
// Use Cases: Archiving financial statements for regulatory retention. | Generating legally compliant invoices in PDF/A‑1a format. | Batch converting workbooks to PDF/A‑1a for a document management system.
// AI Prompts: Show C# code to export a workbook to PDF/A‑1b using Aspose.Cells. | Explain how to embed required metadata for PDF/A‑1a when saving with PdfSaveOptions. | Provide error‑handling patterns for PDF/A‑1a export with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// C# code that creates a Workbook, writes a value to cell A1, configures PdfSaveOptions.Compliance = PdfCompliance.PdfA1a, and saves the workbook as a PDF/A‑1a‑compatible file (PdfA1aOutput.pdf) with Aspose.Cells for .NET.
class PdfA1aExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("PDF/A-1a compliance example");

        // Create PDF save options
        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // Set the compliance level to PDF/A-1a
        saveOptions.Compliance = PdfCompliance.PdfA1a;

        // Save the workbook as a PDF file with the specified compliance level
        workbook.Save("PdfA1aOutput.pdf", saveOptions);
    }
}
