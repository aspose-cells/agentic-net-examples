using System;
using Aspose.Cells;

class PreserveCommentsPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add some data to a cell
        sheet.Cells["A1"].PutValue("Cell with comment");

        // Add a comment to the cell
        int commentIndex = sheet.Comments.Add("A1");
        Comment comment = sheet.Comments[commentIndex];
        comment.Note = "This comment will be retained in the PDF";

        // Configure the worksheet to print comments in place when converting to PDF
        sheet.PageSetup.PrintComments = PrintCommentsType.PrintInPlace;

        // Create PDF save options (optional customizations)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Export document structure can be enabled; it does not affect comments but is shown as an example
            ExportDocumentStructure = true
        };

        // Save the workbook as a PDF with the specified options
        workbook.Save("CommentsPreserved.pdf", pdfOptions);
    }
}