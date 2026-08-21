// Title: Export Workbook to PDF with Cell Background Colors using Aspose.Cells C#
// Description: Demonstrates how to create an Excel workbook, apply solid red and green background styles to cells, configure PdfSaveOptions (ExportDocumentStructure), and save the file as a PDF that retains the original cell shading.
// Keywords: Aspose.Cells PDF export C# | preserve cell shading in PDF | PdfSaveOptions ExportDocumentStructure | Excel to PDF background colors | Aspose.Cells visual fidelity
// Common Searches: Aspose.Cells keep cell colors when converting to PDF | C# export Excel workbook to PDF with background shading | PdfSaveOptions settings for visual consistency | How to retain cell background in PDF using Aspose.Cells
// Developer Intent: Create a PDF from an Excel workbook that maintains the cells' background colors.
// Use Cases: Generating status dashboards where colored cells indicate progress and must appear unchanged in printable PDFs. | Producing invoices that highlight discounted rows with background fills that need to survive PDF conversion. | Designing marketing sheets with brand‑specific cell shading that should be visible in the final PDF document.
// AI Prompts: Provide C# code that uses Aspose.Cells to export a workbook to PDF while preserving cell background colors, including the required PdfSaveOptions configuration. | Explain which PdfSaveOptions properties influence visual fidelity, especially background colors, during Excel‑to‑PDF conversion with Aspose.Cells. | Show a step‑by‑step example of applying solid background styles to cells and saving the workbook as a PDF that keeps those colors intact.

using System;
using System.Drawing;
using Aspose.Cells;

// Demonstrates how to create an Excel workbook, apply solid red and green background styles to cells, configure PdfSaveOptions (ExportDocumentStructure), and save the file as a PDF that retains the original cell shading.
class ExportWorkbookToPdf
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add data to cell A1 and set its background to red
        sheet.Cells["A1"].PutValue("Red Background");
        Style redStyle = workbook.CreateStyle();
        redStyle.ForegroundColor = Color.Red;
        redStyle.Pattern = BackgroundType.Solid;
        sheet.Cells["A1"].SetStyle(redStyle);

        // Add data to cell B2 and set its background to green
        sheet.Cells["B2"].PutValue("Green Background");
        Style greenStyle = workbook.CreateStyle();
        greenStyle.ForegroundColor = Color.Green;
        greenStyle.Pattern = BackgroundType.Solid;
        sheet.Cells["B2"].SetStyle(greenStyle);

        // Configure PDF save options to retain document structure (helps preserve visual layout)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.ExportDocumentStructure = true;

        // Export the workbook to PDF while preserving cell background colors
        workbook.Save("WorkbookWithColors.pdf", pdfOptions);
    }
}
