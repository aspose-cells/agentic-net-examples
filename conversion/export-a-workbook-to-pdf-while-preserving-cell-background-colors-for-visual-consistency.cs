// Title: Export Workbook to PDF with Cell Background Colors using Aspose.Cells (C#)
// Description: Creates a workbook, applies solid red and green background fills to cells, configures PdfSaveOptions, and saves the file as a PDF that keeps the cell colors intact.
// Keywords: Aspose.Cells PDF export | preserve cell background color | C# Excel to PDF | PdfSaveOptions fill color | retain formatting in PDF | solid background fill Aspose | Excel PDF conversion C#
// Common Searches: Aspose.Cells keep cell fill color when converting to PDF | C# export Excel to PDF with background colors | PdfSaveOptions preserve formatting Aspose | how to retain cell background in PDF using Aspose.Cells | export colored Excel cells to PDF C#
// Developer Intent: Export an Excel workbook to PDF while preserving the cells’ background colors.
// Use Cases: Generate printable reports where colored cells highlight status or priority. | Create invoices with colored header rows that must appear in the PDF version. | Produce marketing sheets that use background colors to differentiate product categories.
// AI Prompts: Show how to export multiple worksheets to a single PDF while keeping each sheet’s cell colors. | Demonstrate applying gradient fills to cells and ensuring they render correctly in the PDF output. | Explain which PdfSaveOptions settings improve the visual quality of colored cells in the exported PDF.

using System;
using Aspose.Cells;
using System.Drawing;

// Creates a workbook, applies solid red and green background fills to cells, configures PdfSaveOptions, and saves the file as a PDF that keeps the cell colors intact.
class ExportWorkbookToPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // ---------- Apply background colors ----------
        // Red background for cell A1
        Style redStyle = workbook.CreateStyle();
        redStyle.ForegroundColor = Color.Red;
        redStyle.Pattern = BackgroundType.Solid;
        worksheet.Cells["A1"].PutValue("Red Cell");
        worksheet.Cells["A1"].SetStyle(redStyle);

        // Green background for cell B2
        Style greenStyle = workbook.CreateStyle();
        greenStyle.ForegroundColor = Color.Green;
        greenStyle.Pattern = BackgroundType.Solid;
        worksheet.Cells["B2"].PutValue("Green Cell");
        worksheet.Cells["B2"].SetStyle(greenStyle);

        // ---------- Configure PDF save options ----------
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        // Preserve document structure (optional, does not affect colors)
        pdfOptions.ExportDocumentStructure = true;

        // ---------- Save workbook as PDF ----------
        workbook.Save("WorkbookWithBackgroundColors.pdf", pdfOptions);
    }
}
