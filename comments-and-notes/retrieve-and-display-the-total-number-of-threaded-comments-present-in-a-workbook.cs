using System;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentCount
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the collection of threaded comment authors
            ThreadedCommentAuthorCollection authors = workbook.Worksheets.ThreadedCommentAuthors;
            int authorIdx = authors.Add("John Doe", "john.doe@example.com", "JD");
            ThreadedCommentAuthor author = authors[authorIdx];

            // Add some threaded comments to demonstrate counting
            // Cell A1
            worksheet.Comments.AddThreadedComment("A1", "First comment in A1", author);
            worksheet.Comments.AddThreadedComment("A1", "Reply to first comment", author);

            // Cell B2
            worksheet.Comments.AddThreadedComment("B2", "Only comment in B2", author);

            // Calculate total number of threaded comments in the worksheet
            int totalThreadedComments = 0;
            foreach (Comment comment in worksheet.Comments)
            {
                // Each comment may contain a collection of threaded comments
                totalThreadedComments += comment.ThreadedComments.Count;
            }

            // Display the result
            Console.WriteLine($"Total threaded comments in the workbook: {totalThreadedComments}");

            // Save the workbook (lifecycle rule: save)
            workbook.Save("ThreadedCommentsCountDemo.xlsx");
        }
    }
}