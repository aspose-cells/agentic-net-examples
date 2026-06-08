using System;
using Aspose.Cells;

class RetrieveThreadedCommentCreatedTime
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a threaded comment author to the workbook
        int authorId = workbook.Worksheets.ThreadedCommentAuthors.Add("User1", "user1@example.com", "1");
        ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorId];

        // Add a regular comment to cell A1 (required to hold threaded comments)
        int commentIndex = worksheet.Comments.Add("A1");
        Comment comment = worksheet.Comments[commentIndex];

        // Add a threaded comment to the comment
        int threadedCommentIndex = comment.ThreadedComments.Add("Sample threaded comment", author);
        ThreadedComment threadedComment = comment.ThreadedComments[threadedCommentIndex];

        // Retrieve and log the creation time of the threaded comment
        DateTime createdTime = threadedComment.CreatedTime;
        Console.WriteLine("Threaded comment created at: " + createdTime);

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("ThreadedCommentCreatedTimeDemo.xlsx");
    }
}