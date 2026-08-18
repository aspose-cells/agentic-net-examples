// Title: Export Excel to PDF without cell background colors using Aspose.Cells (C#)
// Description: Demonstrates how to create a workbook, apply a fill color, enable BlackAndWhite printing, and save the result as a minimalist PDF with Aspose.Cells. The BlackAndWhite page‑setup flag strips all cell shading during PDF conversion.
// Keywords: Aspose.Cells PDF export | C# remove cell fill | BlackAndWhite page setup | minimalist PDF from Excel | ignore background colors Aspose | Excel to PDF conversion settings | Aspose.Cells SaveFormat.Pdf
// Common Searches: Aspose.Cells export PDF without background colors | C# convert Excel to black‑and‑white PDF | remove cell shading when saving PDF with Aspose | minimal layout PDF from Excel using Aspose.Cells | how to ignore fill colors in Aspose.Cells PDF output
// Developer Intent: Generate a PDF from an Excel workbook while discarding all cell background colors to achieve a clean, black‑and‑white document.
// Use Cases: Print‑ready reports that need only text and borders, no color distractions. | Corporate invoices or receipts required in a monochrome PDF format. | Archival documentation where a minimalist, color‑free layout is mandated.
// AI Prompts: Provide C# code that uses Aspose.Cells to export an Excel sheet to PDF with the BlackAndWhite property so no cell fills appear. | Show an example of creating a minimalist PDF from a workbook, ignoring background colors, with Aspose.Cells SaveFormat.Pdf. | Explain alternative ways to produce a color‑free PDF in Aspose.Cells, such as configuring PdfSaveOptions.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, apply a fill color, enable BlackAndWhite printing, and save the result as a minimalist PDF with Aspose.Cells. The BlackAndWhite page‑setup flag strips all cell shading during PDF conversion.
class ExportPdfMinimalist
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data
        sheet.Cells["A1"].PutValue("Header");
        sheet.Cells["A2"].PutValue("Item 1");
        sheet.Cells["A3"].PutValue("Item 2");

        // Apply a background color to demonstrate that it will be ignored
        Style bgStyle = workbook.CreateStyle();
        bgStyle.ForegroundColor = Color.Yellow;
        bgStyle.Pattern = BackgroundType.Solid;
        sheet.Cells["A1"].SetStyle(bgStyle);
        sheet.Cells["A2"].SetStyle(bgStyle);
        sheet.Cells["A3"].SetStyle(bgStyle);

        // Enable black‑and‑white printing; this removes cell background colors in the PDF
        sheet.PageSetup.BlackAndWhite = true;

        // Save the workbook as a PDF file
        workbook.Save("MinimalistLayout.pdf", SaveFormat.Pdf);
    }
}
