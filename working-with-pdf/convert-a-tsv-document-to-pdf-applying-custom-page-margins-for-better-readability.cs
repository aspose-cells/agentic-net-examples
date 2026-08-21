// Title: C# – Convert TSV to PDF with Custom Page Margins using Aspose.Cells
// Description: Learn how to load a TSV file into an Aspose.Cells Workbook, set top, bottom, left, and right margins (in centimeters), configure portrait A4 layout, and save the worksheet as a PDF with one page per sheet.
// Keywords: Aspose.Cells TSV to PDF | C# custom page margins | PdfSaveOptions OnePagePerSheet | set margins Aspose.Cells | portrait A4 PDF Aspose | load TSV with LoadOptions | export worksheet to PDF | C# PDF generation Aspose
// Common Searches: how to convert tsv to pdf using aspose.cells c# | set custom margins when saving pdf with aspose.cells | aspose.cells pdfsaveoptions one page per sheet example | change paper size and orientation for pdf output in c# | load tsv file into workbook aspose.cells
// Developer Intent: Load a TSV document, apply specific page margins, and export it as a PDF file.
// Use Cases: Create printable reports from TSV data with precise margin control. | Generate A4 portrait PDFs for batch‑processed TSV files. | Produce one‑page‑per‑sheet PDFs where each worksheet respects custom margins.
// AI Prompts: Show how to set page margins in inches instead of centimeters. | Add a header and footer to the worksheet before exporting to PDF. | Configure PdfSaveOptions to embed fonts while keeping the custom margins.

using System;
using Aspose.Cells;
using System.IO;

// Learn how to load a TSV file into an Aspose.Cells Workbook, set top, bottom, left, and right margins (in centimeters), configure portrait A4 layout, and save the worksheet as a PDF with one page per sheet.
class TsvToPdfWithMargins
{
    static void Main()
    {
        // Input TSV file and output PDF file paths
        string inputTsvPath = "input.tsv";
        string outputPdfPath = "output.pdf";

        // Load the TSV file using LoadOptions to specify the TSV format
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Tsv);
        Workbook workbook = new Workbook(inputTsvPath, loadOptions);

        // Access the first worksheet (the TSV data is loaded into this sheet)
        Worksheet sheet = workbook.Worksheets[0];

        // Set custom page margins (values are in centimeters)
        sheet.PageSetup.TopMargin = 2.0;      // 2 cm top margin
        sheet.PageSetup.BottomMargin = 1.5;   // 1.5 cm bottom margin
        sheet.PageSetup.LeftMargin = 1.0;     // 1 cm left margin
        sheet.PageSetup.RightMargin = 1.0;    // 1 cm right margin

        // Optional: set orientation and paper size for better layout
        sheet.PageSetup.Orientation = PageOrientationType.Portrait;
        sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;

        // Create PDF save options (e.g., one page per sheet)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            OnePagePerSheet = true
        };

        // Save the workbook as a PDF with the specified margins
        workbook.Save(outputPdfPath, pdfOptions);
    }
}
