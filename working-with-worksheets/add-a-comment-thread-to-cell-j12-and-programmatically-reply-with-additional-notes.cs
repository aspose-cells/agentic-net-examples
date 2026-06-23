using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // --------------------------------------------------------------------
        // Create threaded comment authors
        // --------------------------------------------------------------------
        ThreadedCommentAuthorCollection authors = workbook.Worksheets.ThreadedCommentAuthors;
        int aliceIdx = authors.Add("Alice", "alice@example.com", "PROV1");
        ThreadedCommentAuthor alice = authors[aliceIdx];

        int bobIdx = authors.Add("Bob", "bob@example.com", "PROV2");
        ThreadedCommentAuthor bob = authors[bobIdx];

        // --------------------------------------------------------------------
        // Add a threaded comment to cell J12 (row index 11, column index 9)
        // --------------------------------------------------------------------
        worksheet.Comments.AddThreadedComment(11, 9, "Initial comment on J12", alice);

        // --------------------------------------------------------------------
        // Retrieve the threaded comments collection for J12
        // --------------------------------------------------------------------
        ThreadedCommentCollection threadedComments = worksheet.Comments.GetThreadedComments(11, 9);

        // --------------------------------------------------------------------
        // Add a reply (additional note) to the thread
        // --------------------------------------------------------------------
        threadedComments.Add("Reply from Bob", bob);

        // Optionally update the reply's text via the Notes property
        ThreadedComment reply = threadedComments[threadedComments.Count - 1];
        reply.Notes = "Updated reply with additional notes";

        // --------------------------------------------------------------------
        // Save the workbook
        // --------------------------------------------------------------------
        workbook.Save("ThreadedComment_J12.xlsx");
    }
}