// Title: C# – Set PDF version to 1.4 with Aspose.Cells PdfSaveOptions for legacy viewer support
// Description: Shows how to create a workbook, add sample data, set PdfSaveOptions.Compliance to PdfCompliance.Pdf14, and save the workbook as a PDF that conforms to PDF 1.4, guaranteeing compatibility with older PDF readers.
// Keywords: Aspose.Cells | PdfSaveOptions | PDF 1.4 | PdfCompliance.Pdf14 | C# PDF export | legacy PDF viewer compatibility | Excel to PDF 1.4 | set PDF version Aspose | PDF compliance Aspose.Cells
// Common Searches: Aspose.Cells set PDF version 1.4 | PdfSaveOptions compliance PDF 1.4 C# | export Excel to PDF 1.4 using Aspose | C# generate PDF 1.4 from workbook | legacy PDF viewer compatibility Aspose.Cells
// Developer Intent: Configure PDF compliance to version 1.4 when saving a workbook.
// Use Cases: Produce PDF reports that must meet the PDF 1.4 standard for archival or regulatory reasons. | Ensure Excel‑to‑PDF conversions work on older PDF readers that only support version 1.4. | Batch‑process multiple workbooks with a shared PdfSaveOptions instance to enforce PDF 1.4 across all outputs.
// AI Prompts: Give me C# code that sets PdfSaveOptions.Compliance to Pdf14 and saves a workbook as a PDF with Aspose.Cells. | How do I export an Excel worksheet to a PDF file compatible with PDF version 1.4 using Aspose.Cells? | Explain the steps to configure PdfSaveOptions for PDF 1.4 compliance to support legacy PDF viewers.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Shows how to create a workbook, add sample data, set PdfSaveOptions.Compliance to PdfCompliance.Pdf14, and save the workbook as a PDF that conforms to PDF 1.4, guaranteeing compatibility with older PDF readers.
class SetPdfVersionExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add some sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("PDF version 1.4 example");

        // Create PDF save options
        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // Set compliance to PDF 1.4 (Pdf14 or None)
        saveOptions.Compliance = PdfCompliance.Pdf14;

        // Save the workbook as PDF with the specified compliance level
        workbook.Save("OutputPdf14.pdf", saveOptions);
    }
}
