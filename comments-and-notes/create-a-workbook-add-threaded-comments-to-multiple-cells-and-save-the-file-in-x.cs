using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Obtain the collection of threaded comment authors
        ThreadedCommentAuthorCollection authors = workbook.Worksheets.ThreadedCommentAuthors;

        // Add two authors to the workbook
        int aliceIndex = authors.Add("Alice", "alice@example.com", "PROV1");
        ThreadedCommentAuthor alice = authors[aliceIndex];

        int bobIndex = authors.Add("Bob", "bob@example.com", "PROV2");
        ThreadedCommentAuthor bob = authors[bobIndex];

        // Add threaded comments to cell B2 using the cell name overload
        // First comment by Alice
        worksheet.Comments.AddThreadedComment("B2", "Initial comment by Alice.", alice);
        // Reply comment by Bob
        worksheet.Comments.AddThreadedComment("B2", "Reply from Bob.", bob);

        // Add a threaded comment to cell C3 using the row/column overload
        // Row index 2, column index 2 corresponds to cell C3
        worksheet.Comments.AddThreadedComment(2, 2, "Comment on C3 by Alice.", alice);

        // Save the workbook in XLSX format (lifecycle rule: save)
        workbook.Save("ThreadedCommentsMultiple.xlsx");
    }
}