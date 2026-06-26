using System;
using Aspose.Cells;

class RemoveThreadedCommentDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a threaded comment author
        int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add("John Doe", "john.doe@example.com", "PROVIDER");
        ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

        // Add a threaded comment to cell C3
        worksheet.Comments.AddThreadedComment("C3", "Initial threaded comment", author);

        // Retrieve the comment object located at C3
        Comment comment = worksheet.Comments["C3"];

        // Remove the threaded comment using the collection's RemoveAt method (removes the first threaded comment)
        if (comment.ThreadedComments.Count > 0)
        {
            comment.ThreadedComments.RemoveAt(0);
        }

        // Save the workbook
        workbook.Save("RemovedThreadedComment.xlsx");
    }
}