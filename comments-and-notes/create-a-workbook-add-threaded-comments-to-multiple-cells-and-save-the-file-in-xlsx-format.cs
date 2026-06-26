using System;
using Aspose.Cells;

class ThreadedCommentsDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add threaded comment authors
        int authorIdx1 = workbook.Worksheets.ThreadedCommentAuthors.Add("Alice", "alice@example.com", "PROV1");
        ThreadedCommentAuthor author1 = workbook.Worksheets.ThreadedCommentAuthors[authorIdx1];

        int authorIdx2 = workbook.Worksheets.ThreadedCommentAuthors.Add("Bob", "bob@example.com", "PROV2");
        ThreadedCommentAuthor author2 = workbook.Worksheets.ThreadedCommentAuthors[authorIdx2];

        // Add threaded comments to cell B2 using cell name overload
        worksheet.Comments.AddThreadedComment("B2", "First comment by Alice", author1);
        worksheet.Comments.AddThreadedComment("B2", "Reply by Bob", author2);

        // Add threaded comments to cell C3 using row/column overload (row=2, column=2)
        worksheet.Comments.AddThreadedComment(2, 2, "Another comment on C3 by Alice", author1);
        worksheet.Comments.AddThreadedComment(2, 2, "Follow‑up on C3 by Bob", author2);

        // Save the workbook in XLSX format
        workbook.Save("ThreadedCommentsDemo.xlsx");
    }
}