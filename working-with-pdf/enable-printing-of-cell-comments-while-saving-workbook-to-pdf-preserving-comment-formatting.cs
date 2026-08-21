// Title: Print Cell Comments with Original Formatting to PDF using Aspose.Cells (C#)
// Description: Demonstrates how to add a comment to a worksheet, configure PageSetup to print comments in place, and save the workbook as a PDF with the comment rendered exactly as it appears in Excel. Uses default PdfSaveOptions for a quick conversion.
// Keywords: Aspose.Cells PDF comments | C# print cell comments to PDF | PrintCommentsType.PrintInPlace example | save workbook with comments Aspose.Cells | preserve comment appearance PDF
// Common Searches: include cell comments when exporting Aspose.Cells workbook to PDF C# | Aspose.Cells keep comment formatting in PDF output | how to print comments in place with Aspose.Cells | export Excel comments to PDF using Aspose.Cells
// Developer Intent: Generate a PDF from an Aspose.Cells workbook that displays cell comments with their original styling.
// Use Cases: Create printable reports that show reviewer notes stored as cell comments. | Distribute audit‑trail spreadsheets as PDFs while retaining comment visuals. | Produce documentation where inline comments appear alongside data cells.
// AI Prompts: Show how to limit printed comments to the first page of the PDF with Aspose.Cells. | Give an example of changing a comment's background color before saving to PDF. | Explain how to set page orientation in PdfSaveOptions while keeping comments printed.

using System;
using Aspose.Cells;

// Demonstrates how to add a comment to a worksheet, configure PageSetup to print comments in place, and save the workbook as a PDF with the comment rendered exactly as it appears in Excel. Uses default PdfSaveOptions for a quick conversion.
class PrintCommentsToPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a comment to cell A1
        int commentIndex = sheet.Comments.Add("A1");
        Comment comment = sheet.Comments[commentIndex];
        comment.Note = "This is a test comment";

        // Set the page setup to print comments as they appear on the sheet
        // This preserves the comment formatting in the PDF output
        sheet.PageSetup.PrintComments = PrintCommentsType.PrintInPlace;

        // Create PDF save options (default options are sufficient for comments)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the workbook to PDF with comments printed
        workbook.Save("CommentsPrinted.pdf", pdfOptions);
    }
}
