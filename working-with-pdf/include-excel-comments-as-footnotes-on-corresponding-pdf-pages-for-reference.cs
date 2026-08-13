// Title: C# – Export Excel Cell Comments as PDF Footnotes with Aspose.Cells
// Description: Shows how to build a workbook, add cell comments, set PrintCommentsType to PrintSheetEnd, add a right‑aligned page‑number footer, enable document‑structure export, and save the sheet as a PDF where each comment appears as a footnote on the corresponding page.
// Keywords: Aspose.Cells C# PDF export | Excel comments footnotes PDF | PrintCommentsType PrintSheetEnd | PdfSaveOptions ExportDocumentStructure | add footer page number Aspose.Cells | convert Excel to PDF with comments | Aspose.Cells comment footnote example | C# workbook to PDF with footnotes
// Common Searches: Aspose.Cells export comments as footnotes | C# print Excel comments at end of sheet PDF | how to add page numbers when saving Excel to PDF with Aspose | enable document structure in PDF using Aspose.Cells | convert Excel to PDF with comment footnotes C#
// Developer Intent: Create a PDF from an Excel workbook where each cell comment is rendered as a footnote on the same page, optionally with a page‑number footer.
// Use Cases: Product catalog where item notes appear as footnotes in the PDF | Regulatory compliance report with explanatory comments displayed as footnotes | Financial statements that need cell annotations and page numbers in the exported PDF | Academic worksheets where teacher comments are shown as footnotes | Invoices that include comment‑based terms and conditions as footnotes
// AI Prompts: Generate C# code using Aspose.Cells to convert an Excel sheet to PDF with cell comments shown as footnotes and a page‑number footer. | Show how to set PrintCommentsType to PrintSheetEnd and enable ExportDocumentStructure in PdfSaveOptions. | Provide an example that adds multiple comments, configures footnote printing, and saves the workbook as a PDF using Aspose.Cells for .NET. | Explain how to include Excel comments as end‑of‑sheet footnotes when exporting to PDF with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsCommentFootnotesPdf
{
    // Shows how to build a workbook, add cell comments, set PrintCommentsType to PrintSheetEnd, add a right‑aligned page‑number footer, enable document‑structure export, and save the sheet as a PDF where each comment appears as a footnote on the corresponding page.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1.20);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(0.80);

            // Add comments to cells that will appear as footnotes
            int commentIdx = sheet.Comments.Add("A2");
            Comment comment = sheet.Comments[commentIdx];
            comment.Note = "Fresh apples from the orchard.";

            commentIdx = sheet.Comments.Add("A3");
            comment = sheet.Comments[commentIdx];
            comment.Note = "Ripe bananas imported from Ecuador.";

            // Configure the worksheet to print comments at the end of the sheet
            sheet.PageSetup.PrintComments = PrintCommentsType.PrintSheetEnd;

            // Optional: add a footer with page number for reference
            sheet.PageSetup.SetFooter(2, "&P of &N"); // Right section

            // Set PDF save options (e.g., export document structure)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true
            };

            // Save the workbook as PDF; comments will appear as footnotes
            workbook.Save("CommentsFootnotes.pdf", pdfOptions);
        }
    }
}
