// Title: Convert Excel Cell Comments to PDF Annotations with Aspose.Cells for .NET
// Description: Demonstrates how to save a workbook as a PDF while keeping cell comments as interactive PDF annotations by enabling ExportDocumentStructure in PdfSaveOptions.
// Keywords: Aspose.Cells | C# PDF conversion | Excel comments to PDF | ExportDocumentStructure | PDF annotations | preserve cell comments | PdfSaveOptions | Aspose.Cells for .NET | document structure export | batch Excel to PDF
// Common Searches: keep Excel comments when converting to PDF C# | Aspose.Cells ExportDocumentStructure example | PDF annotation from Excel comment .NET | save workbook as PDF with comments Aspose | how to export cell notes to PDF using Aspose.Cells
// Developer Intent: Create a PDF from an Excel workbook that retains each cell's comment as a clickable annotation.
// Use Cases: Generate review‑ready PDFs where stakeholder notes appear as annotation pop‑ups. | Automate compliance‑driven batch conversions that must preserve comment metadata. | Distribute printable reports that include explanatory remarks without sharing the original spreadsheet.
// AI Prompts: Show how to export only selected comments as PDF annotations with Aspose.Cells. | Explain ways to style comment annotations (color, font, icon) in the resulting PDF. | Provide a script to process a folder of workbooks, converting each to PDF while keeping all comments.

using System;
using Aspose.Cells;

// Demonstrates how to save a workbook as a PDF while keeping cell comments as interactive PDF annotations by enabling ExportDocumentStructure in PdfSaveOptions.
class ExportCommentsToPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data
        worksheet.Cells["A1"].PutValue("Hello");
        worksheet.Cells["B1"].PutValue("World");

        // Add a comment to cell A1
        // The comment will be exported as a PDF annotation when document structure is exported
        int commentIndex = worksheet.Comments.Add("A1");
        Comment comment = worksheet.Comments[commentIndex];
        comment.Note = "This is a sample comment";

        // Configure PDF save options
        // ExportDocumentStructure = true ensures that comments are retained as annotations
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        pdfSaveOptions.ExportDocumentStructure = true;

        // Save the workbook as a PDF file using the configured options
        workbook.Save("CommentsExported.pdf", pdfSaveOptions);
    }
}
