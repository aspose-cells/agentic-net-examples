// Title: Create PDF/A‑1a compliant PDF from Excel using Aspose.Cells for .NET (C#)
// Description: Shows how to generate an archival PDF/A‑1a file from an Excel workbook with Aspose.Cells for .NET by configuring PdfSaveOptions.Compliance to PdfCompliance.PdfA1a and saving the workbook.
// Keywords: Aspose.Cells | PdfSaveOptions | PdfCompliance.PdfA1a | PDF/A-1a | C# export Excel to PDF | archival PDF | long‑term preservation | regulatory compliance | Excel to PDF/A | .NET
// Common Searches: Aspose.Cells export Excel to PDF/A-1a C# | set PdfSaveOptions compliance PDF/A-1a | how to create PDF/A-1a with Aspose.Cells | C# code for PDF/A-1a compliance | generate archival PDF from workbook Aspose
// Developer Intent: Generate a PDF/A‑1a compliant PDF from an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Preserve financial statements in PDF/A‑1a for long‑term archiving. | Submit regulatory reports in PDF/A‑1a format to meet compliance standards. | Automate PDF/A‑1a conversion within a document‑management workflow. | Create legally binding PDFs from spreadsheets for contracts or audits.
// AI Prompts: Show C# code to set PdfSaveOptions.Compliance to PdfA1b for a different PDF/A level. | Provide an example that adds custom XMP metadata to a PDF/A‑1a file saved with Aspose.Cells. | Explain how to batch‑convert a list of workbooks to PDF/A‑1a using a loop. | How can I verify PDF/A‑1a compliance after saving a workbook with Aspose.Cells?

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfComplianceDemo
{
    // Shows how to generate an archival PDF/A‑1a file from an Excel workbook with Aspose.Cells for .NET by configuring PdfSaveOptions.Compliance to PdfCompliance.PdfA1a and saving the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Add some sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("PDF/A‑1a compliance example");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Create PDF save options and set compliance to PDF/A‑1a
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA1a
            };

            // Save the workbook as a PDF with the specified compliance level (lifecycle: save)
            workbook.Save("PdfA1aOutput.pdf", pdfOptions);

            Console.WriteLine("PDF saved with PDF/A‑1a compliance.");
        }
    }
}
