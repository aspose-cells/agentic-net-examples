// Title: C# – Convert an Excel workbook to PDF with cell comments displayed as footnotes using Aspose.Cells
// Description: Shows how to add comments to cells, set PageSetup.PrintComments to PrintSheetEnd, enable ExportDocumentStructure in PdfSaveOptions, and save the workbook as a PDF where every comment appears as a footnote at the end of the sheet.
// Keywords: Aspose.Cells | C# | Excel to PDF | cell comments | footnotes | PrintCommentsType | PrintSheetEnd | PdfSaveOptions | ExportDocumentStructure | include comments in PDF
// Common Searches: Aspose.Cells export comments as footnotes PDF | C# save Excel as PDF with comments at end of sheet | PrintCommentsType PrintSheetEnd example | PdfSaveOptions ExportDocumentStructure usage | how to include Excel cell comments in PDF
// Developer Intent: Generate a PDF from an Excel workbook where all cell comments are rendered as footnotes.
// Use Cases: Create a sales‑report workbook, add explanatory notes to product cells, set PrintComments to PrintSheetEnd, and export to PDF so reviewers see the notes as footnotes. | Export a financial model with analyst annotations, preserving the comments as end‑of‑sheet footnotes by enabling ExportDocumentStructure in the PDF options. | Produce printable documentation from a spreadsheet that contains instructional comments, automatically converting those comments into footnotes in the final PDF.
// AI Prompts: Write C# code with Aspose.Cells that saves a workbook to PDF and renders every cell comment as a footnote at the end of the sheet. | Explain how to configure PageSetup.PrintComments = PrintSheetEnd and PdfSaveOptions.ExportDocumentStructure for PDF conversion with comments. | Provide guidance on customizing the appearance of comment footnotes when exporting an Excel worksheet to PDF using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsCommentToPdf
{
    // Shows how to add comments to cells, set PageSetup.PrintComments to PrintSheetEnd, enable ExportDocumentStructure in PdfSaveOptions, and save the workbook as a PDF where every comment appears as a footnote at the end of the sheet.
    class Program
    {
        static void Main()
        {
            // ---------- Create ----------
            // Initialize a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Bananas");
            sheet.Cells["B3"].PutValue(85);

            // Add comments to cells
            int commentIdx = sheet.Comments.Add("A2");
            Comment comment = sheet.Comments[commentIdx];
            comment.Note = "Seasonal fruit";

            commentIdx = sheet.Comments.Add("B3");
            comment = sheet.Comments[commentIdx];
            comment.Note = "Estimated demand";

            // ---------- Configure printing of comments ----------
            // Print comments at the end of the sheet so they appear as footnotes in PDF
            sheet.PageSetup.PrintComments = PrintCommentsType.PrintSheetEnd;

            // ---------- Save ----------
            // Set PDF save options (optional: export document structure)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true
            };

            // Save the workbook as PDF; comments will be rendered as footnotes
            workbook.Save("WorkbookWithComments.pdf", pdfOptions);

            Console.WriteLine("Workbook successfully saved to PDF with comments as footnotes.");
        }
    }
}
