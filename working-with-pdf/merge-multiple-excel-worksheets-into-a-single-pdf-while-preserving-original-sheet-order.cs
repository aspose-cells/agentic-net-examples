// Title: C# – Aspose.Cells: Merge All Excel Worksheets into One PDF (Preserve Sheet Order)
// Description: Loads an Excel workbook with Aspose.Cells, sets PdfSaveOptions.SheetSet to All, and saves the entire workbook as a single PDF file while keeping the worksheets in their original sequence.
// Keywords: Aspose.Cells PDF export | C# merge Excel sheets to PDF | preserve sheet order PDF | PdfSaveOptions SheetSet.All example | convert multi‑sheet workbook to PDF
// Common Searches: export all Excel worksheets to one PDF Aspose.Cells | keep original sheet order when converting Excel to PDF .NET | C# combine multiple sheets into single PDF | Aspose.Cells PDFSaveOptions for whole workbook
// Developer Intent: Generate one PDF that contains every worksheet from an Excel file in the same order as the source workbook.
// Use Cases: Create a unified report PDF from a financial workbook with several tabs. | Archive project spreadsheets as a single printable document for stakeholders. | Distribute a complete client proposal where all Excel sheets are bundled into one PDF.
// AI Prompts: Show how to export only selected sheets to PDF while retaining their order with Aspose.Cells. | Explain how to assign different page orientations to individual worksheets when merging them into one PDF. | Provide code to prepend a custom cover page before combining Excel sheets into a PDF using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;   // for SheetSet and PdfSaveOptions

// Loads an Excel workbook with Aspose.Cells, sets PdfSaveOptions.SheetSet to All, and saves the entire workbook as a single PDF file while keeping the worksheets in their original sequence.
class MergeSheetsToPdf
{
    static void Main()
    {
        // Path to the source Excel file containing multiple worksheets
        string excelPath = "input.xlsx";

        // Load the workbook (preserves original sheet order)
        Workbook workbook = new Workbook(excelPath);

        // Prepare PDF save options to include all sheets in their original order
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            SheetSet = SheetSet.All   // ensures every worksheet is rendered
        };

        // Path for the resulting PDF file
        string pdfPath = "merged_output.pdf";

        // Save the workbook as a single PDF document
        workbook.Save(pdfPath, pdfOptions);

        Console.WriteLine($"Workbook successfully merged into PDF: {pdfPath}");
    }
}
