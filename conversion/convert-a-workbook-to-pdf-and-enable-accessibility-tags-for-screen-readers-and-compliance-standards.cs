// Title: Convert an Aspose.Cells Workbook to PDF with Accessibility Tags and PDF/A‑1b Compliance (C#)
// Description: Demonstrates how to create or load a workbook, set PdfSaveOptions to export document structure for screen‑reader accessibility, apply PDF/A‑1b compliance, and save the result as an accessible PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells PDF accessibility | ExportDocumentStructure | PdfSaveOptions | PDF/A‑1b C# | accessible PDF from Excel | screen reader tags | WCAG compliance Aspose | ADA PDF export | Aspose.Cells conversion C#
// Common Searches: Aspose.Cells enable accessibility tags when saving PDF | C# set PDF/A‑1b compliance with Aspose.Cells | how to export document structure for screen readers in Aspose.Cells | convert Excel workbook to accessible PDF using Aspose | Aspose.Cells PDF/A‑1b export example
// Developer Intent: Produce a PDF from an Excel workbook that includes PDF/UA‑compatible tags and meets PDF/A‑1b archival standards using Aspose.Cells for .NET.
// Use Cases: Generate PDF reports that are readable by screen‑reader software for visually impaired users. | Archive financial or regulatory spreadsheets in PDF/A‑1b format for long‑term preservation. | Automate creation of accessible PDFs to satisfy WCAG, ADA, or Section 508 compliance.
// AI Prompts: Write C# code with Aspose.Cells to convert an existing Excel file to a PDF that has ExportDocumentStructure enabled and PDF/A‑1b compliance. | Explain the impact of ExportDocumentStructure and PdfCompliance settings on PDF accessibility in Aspose.Cells. | Provide a step‑by‑step tutorial for adding PDF/UA tags to a workbook saved as PDF with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;   // for PdfCompliance enum

namespace AsposeCellsPdfAccessibilityDemo
{
    // Demonstrates how to create or load a workbook, set PdfSaveOptions to export document structure for screen‑reader accessibility, apply PDF/A‑1b compliance, and save the result as an accessible PDF using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a new workbook (or load an existing file)
            // -------------------------------------------------
            // Using the provided constructor rule for Workbook creation.
            Workbook workbook = new Workbook();

            // Add sample data – this step is optional if you load an existing file.
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Employee");
            sheet.Cells["B1"].PutValue("Salary");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(50000);
            sheet.Cells["A3"].PutValue("Alice");
            sheet.Cells["B3"].PutValue(62000);

            // -------------------------------------------------
            // 2. Configure PDF save options for accessibility
            // -------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Enable export of document structure (tags) for screen readers.
            pdfOptions.ExportDocumentStructure = true;

            // Set PDF/A compliance (e.g., PDF/A-1b) to meet archival standards.
            pdfOptions.Compliance = PdfCompliance.PdfA1b;

            // -------------------------------------------------
            // 3. Save the workbook as a PDF using the provided Save rule
            // -------------------------------------------------
            // This uses the Workbook.Save(string, SaveOptions) overload.
            workbook.Save("AccessibleDocument.pdf", pdfOptions);

            Console.WriteLine("Workbook successfully saved as PDF with accessibility tags.");
        }
    }
}
