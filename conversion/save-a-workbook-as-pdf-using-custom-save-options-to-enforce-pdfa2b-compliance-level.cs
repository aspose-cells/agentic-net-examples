// Title: Export Excel to PDF/A‑2b with Aspose.Cells C# PdfSaveOptions
// Description: Creates a workbook, configures PdfSaveOptions.Compliance to PdfA2b, and saves the result as a PDF/A‑2b compliant document.
// Keywords: Aspose.Cells PDF/A‑2b export | C# PdfSaveOptions compliance | Excel to PDF/A‑2b conversion | PdfA2b Aspose.Cells example | archival PDF generation C#
// Common Searches: Aspose.Cells export Excel to PDF/A‑2b C# | set PDF compliance level PdfA2b Aspose | C# code for PDF/A‑2b compliant workbook | how to create PDF/A‑2b file with Aspose.Cells | PdfSaveOptions compliance property usage
// Developer Intent: Generate a PDF/A‑2b archival‑ready file from an Excel workbook using Aspose.Cells in C#.
// Use Cases: Long‑term storage of financial statements | Regulatory‑compliant invoice archiving | Batch conversion of reports for legal submission
// AI Prompts: Show a C# snippet that saves an Aspose.Cells workbook as PDF/A‑2b and adds custom document metadata. | Explain how to embed fonts and set PDF/A‑2b conformance options with PdfSaveOptions. | Create code that converts a workbook to PDF/A‑2b and uploads the output to Azure Blob Storage.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Creates a workbook, configures PdfSaveOptions.Compliance to PdfA2b, and saves the result as a PDF/A‑2b compliant document.
class PdfA2bSaveExample
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("PDF/A-2b compliance example");

        // Initialize PDF save options and set the compliance level to PDF/A-2b
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.Compliance = PdfCompliance.PdfA2b;

        // Save the workbook as a PDF file using the specified options
        workbook.Save("PdfA2bOutput.pdf", saveOptions);
    }
}
