using System;
using Aspose.Cells;

namespace ThreadedCommentsCounter
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a threaded comment author
            int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add("John Doe", "john.doe@example.com", "JD");
            ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

            // Add some threaded comments to cell B2
            worksheet.Comments.AddThreadedComment("B2", "Initial comment", author);
            worksheet.Comments.AddThreadedComment("B2", "Reply to initial comment", author);

            // Add another threaded comment to cell C3
            worksheet.Comments.AddThreadedComment("C3", "Another comment", author);

            // Calculate total number of threaded comments in the workbook
            int totalThreadedComments = 0;
            foreach (Worksheet ws in workbook.Worksheets)
            {
                CommentCollection comments = ws.Comments;
                foreach (Comment comment in comments)
                {
                    // Each comment may contain a collection of threaded comments
                    totalThreadedComments += comment.ThreadedComments.Count;
                }
            }

            // Display the result
            Console.WriteLine($"Total threaded comments in the workbook: {totalThreadedComments}");

            // Save the workbook (optional, just to demonstrate the save lifecycle)
            workbook.Save("ThreadedCommentsCount.xlsx");
        }
    }
}