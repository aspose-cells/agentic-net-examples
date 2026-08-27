// Title: Export an Aspose.Cells workbook to a minimalist PDF by disabling cell background colors (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to save a workbook as PDF while setting PageSetup.BlackAndWhite to true. | Show how to configure PdfSaveOptions to produce a PDF without the document structure and without cell fill colors. | Provide a complete example that creates a styled cell, enables black‑and‑white printing, and exports the sheet to a clean PDF file.
// Common Searches: Aspose.Cells C# export Excel to PDF without cell fill colors | How to create a black and white PDF from a worksheet using Aspose.Cells | Minimalist PDF output from Aspose.Cells ignoring background colors | PdfSaveOptions settings to remove background shading in Aspose.Cells PDF conversion | PageSetup.BlackAndWhite effect on PDF export in Aspose.Cells .NET
// Tags: Aspose.Cells PDF export omit cell fills | C# PageSetup.BlackAndWhite for PDF | PdfSaveOptions disable document structure | minimalist PDF layout Aspose.Cells | ignore worksheet background colors PDF

using System;
using System.Drawing;
using Aspose.Cells;

// Creates a workbook, adds sample data, applies a yellow fill to a cell, enables PageSetup.BlackAndWhite to suppress background colors, configures PdfSaveOptions (ExportDocumentStructure = false), and saves the result as MinimalistLayout.pdf, producing a clean black‑and‑white PDF.
class ExportPdfMinimalist
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data
        worksheet.Cells["A1"].PutValue("Header");
        worksheet.Cells["A2"].PutValue("Item 1");
        worksheet.Cells["A3"].PutValue("Item 2");

        // Apply a background color to demonstrate that it will be ignored in the PDF
        Style bgStyle = workbook.CreateStyle();
        bgStyle.ForegroundColor = Color.Yellow;
        bgStyle.Pattern = BackgroundType.Solid;
        worksheet.Cells["A1"].SetStyle(bgStyle);

        // Set the page to print in black and white.
        // This causes background colors to be omitted, yielding a minimalist layout.
        worksheet.PageSetup.BlackAndWhite = true;

        // Create PDF save options (optional customizations can be added here)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Example: do not export the document structure
            ExportDocumentStructure = false
        };

        // Save the workbook as a PDF using the specified options
        workbook.Save("MinimalistLayout.pdf", pdfOptions);
    }
}
