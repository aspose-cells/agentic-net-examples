using System;
using Aspose.Cells;

namespace ThreadedCommentCopyDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            CommentCollection comments = worksheet.Comments;

            // ------------------------------------------------------------
            // Ensure there is a threaded comment in cell E5 for demonstration
            // ------------------------------------------------------------
            int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add(
                "Demo Author",          // Name
                "demoUser@example.com", // UserId
                "DemoProvider");        // ProviderId

            ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

            // Add a threaded comment to cell E5
            comments.AddThreadedComment("E5", "Original comment text", author);

            // ------------------------------------------------------------
            // Retrieve all threaded comments from cell E5
            // ------------------------------------------------------------
            ThreadedCommentCollection sourceComments = comments.GetThreadedComments("E5");

            // ------------------------------------------------------------
            // Copy each threaded comment to cell F6, preserving author and text
            // ------------------------------------------------------------
            foreach (ThreadedComment srcComment in sourceComments)
            {
                // Add a new threaded comment to F6 with the same notes and author
                comments.AddThreadedComment("F6", srcComment.Notes, srcComment.Author);
            }

            // ------------------------------------------------------------
            // Save the workbook
            // ------------------------------------------------------------
            workbook.Save("CopyThreadedComment.xlsx");
        }
    }
}