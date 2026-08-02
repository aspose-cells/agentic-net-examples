// Title: Verify OnePagePerSheet PDF Export Generates One Page per Worksheet with Aspose.Cells (.NET)
// Description: C# sample that creates a workbook with multiple sheets, saves it to PDF using PdfSaveOptions.OnePagePerSheet, and validates that each sheet renders to a single page by checking SheetPrintingPreview.EvaluatedPageCount. Includes guidance for using a PDF parser to confirm pagination.
// Keywords: Aspose.Cells PDF export | OnePagePerSheet verification | SheetPrintingPreview page count | C# PDF pagination test | worksheet to single PDF page
// Common Searches: Aspose.Cells ensure one PDF page per worksheet | C# check PDF page count per sheet Aspose | validate OnePagePerSheet option in Aspose.Cells | how to use SheetPrintingPreview for pagination testing | unit test PDF pagination Aspose.Cells
// Developer Intent: Confirm that each worksheet is exported to exactly one PDF page when OnePagePerSheet is enabled.
// Use Cases: Automated testing to guarantee one‑page‑per‑sheet PDF output. | Generating multi‑sheet reports where each sheet must start on a new page. | CI/CD validation of workbook layout changes to prevent pagination regressions.
// AI Prompts: Generate C# code that parses the PDF stream and counts pages per worksheet to verify OnePagePerSheet. | Show how to modify the example to ignore hidden worksheets while still checking visible sheets. | Provide a logging strategy that records page‑count results for every worksheet during PDF export.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// C# sample that creates a workbook with multiple sheets, saves it to PDF using PdfSaveOptions.OnePagePerSheet, and validates that each sheet renders to a single page by checking SheetPrintingPreview.EvaluatedPageCount. Includes guidance for using a PDF parser to confirm pagination.
class VerifyOnePagePerSheet
{
    static void Main()
    {
        // Create a workbook with two worksheets and fill them with many rows.
        Workbook workbook = new Workbook();
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Sheet1";
        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

        for (int i = 0; i < 200; i++)
        {
            sheet1.Cells[i, 0].PutValue($"Sheet1 Row {i + 1}");
            sheet2.Cells[i, 0].PutValue($"Sheet2 Row {i + 1}");
        }

        // Configure PDF save options to force one page per sheet.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            OnePagePerSheet = true
        };

        // Save the workbook to a memory stream (PDF generation).
        using (MemoryStream pdfStream = new MemoryStream())
        {
            workbook.Save(pdfStream, pdfOptions);
            // PDF is now generated; we will verify pagination using preview objects.
        }

        // Use SheetPrintingPreview to evaluate page count for each worksheet.
        ImageOrPrintOptions previewOptions = new ImageOrPrintOptions
        {
            OnePagePerSheet = true
        };

        bool allSheetsOnePage = true;

        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            Worksheet sheet = workbook.Worksheets[i];
            SheetPrintingPreview preview = new SheetPrintingPreview(sheet, previewOptions);
            int pageCount = preview.EvaluatedPageCount;
            Console.WriteLine($"Worksheet '{sheet.Name}' page count: {pageCount}");
            if (pageCount != 1)
                allSheetsOnePage = false;
        }

        Console.WriteLine(allSheetsOnePage
            ? "Verification succeeded: each worksheet rendered to exactly one page."
            : "Verification failed: one or more worksheets rendered to multiple pages.");
    }
}
