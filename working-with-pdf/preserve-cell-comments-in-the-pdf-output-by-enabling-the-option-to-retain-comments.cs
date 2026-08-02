// Title: Preserve Excel cell comments in PDF using Aspose.Cells for .NET
// Description: Demonstrates how to add a visible comment to a worksheet cell, configure the page setup with PrintCommentsType.PrintInPlace, and save the workbook as a PDF via PdfSaveOptions so the comment is retained in the generated document.
// Keywords: Aspose.Cells | C# | .NET | PDF export | preserve comments | PrintCommentsType.PrintInPlace | cell annotation | Workbook.Save PDF | PdfSaveOptions | Excel to PDF conversion
// Common Searches: Aspose.Cells keep cell comments in PDF | Export Excel comments to PDF C# | PrintCommentsType.PrintInPlace example | Save workbook as PDF with comments visible | Aspose.Cells PDF comment retention .NET
// Developer Intent: Export an Excel worksheet to PDF while retaining visible cell comments.
// Use Cases: Create PDF reports that include reviewer notes attached to specific cells | Generate printable financial statements with audit comments preserved | Distribute spreadsheet documentation where annotations must appear in the PDF
// AI Prompts: Show how to export only visible comments to PDF with Aspose.Cells. | Provide code to hide comments in the PDF while keeping them in the workbook. | Explain how to customize comment font and color when printing to PDF using Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to add a visible comment to a worksheet cell, configure the page setup with PrintCommentsType.PrintInPlace, and save the workbook as a PDF via PdfSaveOptions so the comment is retained in the generated document.
class PreserveCommentsPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add some data to a cell
        worksheet.Cells["A1"].PutValue("Sample data with comment");

        // Add a comment to the cell A1
        int commentIndex = worksheet.Comments.Add("A1");
        Comment comment = worksheet.Comments[commentIndex];
        comment.Note = "This comment will be retained in the PDF output.";
        comment.IsVisible = true; // make sure the comment is visible

        // Set the page setup to print comments in place (so they appear in the PDF)
        worksheet.PageSetup.PrintComments = PrintCommentsType.PrintInPlace;

        // Create PDF save options (default settings)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the workbook as a PDF file with comments preserved
        workbook.Save("CommentsPreserved.pdf", pdfOptions);
    }
}
