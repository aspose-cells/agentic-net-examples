// Title: Aspose.Cells C# – Convert Excel to PDF with Accessibility Tags Disabled
// Description: Load an .xlsx file with Aspose.Cells, evaluate all formulas, set PdfSaveOptions.ExportDocumentStructure to false, and save the workbook as a PDF. The output is a lightweight PDF without PDF/UA accessibility tags, ideal for legacy viewers or batch processing.
// Keywords: Aspose.Cells PDF conversion C# | ExportDocumentStructure false | disable PDF accessibility tags | Excel to PDF without tags | PdfSaveOptions | calculate formulas Aspose.Cells | C# workbook to PDF | remove document structure PDF | lightweight PDF generation | batch PDF export Aspose
// Common Searches: Aspose.Cells disable accessibility tags when saving PDF | PdfSaveOptions ExportDocumentStructure C# example | Convert Excel to PDF without PDF/UA tags Aspose | How to turn off document structure in Aspose.Cells PDF export | C# code to export workbook to PDF without tags
// Developer Intent: Generate a PDF from an Excel workbook while suppressing PDF accessibility tags to produce a simpler, smaller file.
// Use Cases: Printing financial reports where PDF tags are unnecessary | Batch converting spreadsheets to PDFs for archival systems lacking PDF/UA support | Creating lightweight PDFs for mobile devices or email attachments | Ensuring all formula results are up‑to‑date before conversion | Producing PDFs for third‑party tools that require plain document structure
// AI Prompts: Write C# code using Aspose.Cells to convert an Excel file to PDF and disable ExportDocumentStructure. | Explain why setting ExportDocumentStructure to false removes accessibility tags in the resulting PDF. | Show how to evaluate all formulas in a workbook before saving it as PDF with Aspose.Cells. | Provide a step‑by‑step guide to generate a lightweight PDF from Excel using PdfSaveOptions.

using System;
using Aspose.Cells;

// Load an .xlsx file with Aspose.Cells, evaluate all formulas, set PdfSaveOptions.ExportDocumentStructure to false, and save the workbook as a PDF. The output is a lightweight PDF without PDF/UA accessibility tags, ideal for legacy viewers or batch processing.
class Program
{
    static void Main()
    {
        // Load the source Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create PDF save options (using the provided PdfSaveOptions constructor)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Disable exporting of document structure (accessibility tags) for a simpler PDF
        pdfOptions.ExportDocumentStructure = false;

        // Ensure formulas are evaluated before conversion
        workbook.CalculateFormula();

        // Save the workbook as a PDF with the specified options
        workbook.Save("output.pdf", pdfOptions);
    }
}
