// Title: C# – Remove All Pivot Tables and Export Workbook to PDF with Aspose.Cells
// Description: Loads an Excel file, clears every pivot table on all worksheets using Workbook.Worksheets.ClearPivottables(), sets PdfSaveOptions (including IgnoreError), and saves the cleaned workbook as a PDF document.
// Keywords: Aspose.Cells clear pivot tables | C# remove pivot tables Excel | export workbook to PDF Aspose.Cells | .NET PDF save options | IgnoreError PDF export | clean PDF report from Excel | Aspose.Cells pivot table removal
// Common Searches: how to delete all pivot tables before PDF export using Aspose.Cells | Aspose.Cells .NET remove pivot tables and save as PDF | PdfSaveOptions.IgnoreError effect on Excel to PDF conversion | clear pivot tables programmatically C# Aspose | generate clean PDF report from Excel with Aspose.Cells
// Developer Intent: Strip every pivot table from an Excel workbook and generate a PDF without those elements.
// Use Cases: Produce regulatory‑compliant PDFs by eliminating pivot tables that may cause rendering issues. | Batch‑process multiple workbooks to remove pivot tables before archiving them as PDFs. | Create a clean, printable report from a template workbook that contains hidden pivot tables.
// AI Prompts: Write C# code that removes pivot tables only from selected worksheets and saves the file as a PDF with custom page margins using Aspose.Cells. | Explain the role of PdfSaveOptions.IgnoreError when exporting Excel files that contain unsupported pivot features. | Show how to iterate through worksheets, delete specific pivot tables, and then export the workbook to a password‑protected PDF.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Loads an Excel file, clears every pivot table on all worksheets using Workbook.Worksheets.ClearPivottables(), sets PdfSaveOptions (including IgnoreError), and saves the cleaned workbook as a PDF document.
class CleanPdfReport
{
    static void Main()
    {
        // Load the workbook that contains pivot tables
        Workbook workbook = new Workbook("input.xlsx");

        // Remove all pivot tables from every worksheet
        workbook.Worksheets.ClearPivottables();

        // Configure PDF save options (optional: ignore rendering errors)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.IgnoreError = true;

        // Save the cleaned workbook as a PDF file
        workbook.Save("CleanReport.pdf", pdfOptions);
    }
}
