using System;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a threaded comment author
            int authorIdx = workbook.Worksheets.ThreadedCommentAuthors.Add("John Doe", "john.doe@example.com", "Provider1");
            ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIdx];

            // Add a regular comment to cell A1 (required to hold threaded comments)
            int commentIdx = worksheet.Comments.Add("A1");
            Comment comment = worksheet.Comments[commentIdx];

            // Add several threaded comments to the comment
            comment.ThreadedComments.Add("First threaded comment", author);
            comment.ThreadedComments.Add("Second threaded comment", author);
            comment.ThreadedComments.Add("Third threaded comment", author);

            // Retrieve the collection of threaded comments
            ThreadedCommentCollection threadedComments = comment.ThreadedComments;

            // Iterate through the collection and display details
            foreach (ThreadedComment tc in threadedComments)
            {
                // Text of the comment
                string text = tc.Notes;

                // Author name
                string authorName = tc.Author != null ? tc.Author.Name : "Unknown";

                // Creation timestamp (if not set, Aspose.Cells may provide default DateTime)
                DateTime created = tc.CreatedTime;

                Console.WriteLine($"Comment: \"{text}\"");
                Console.WriteLine($"Author : {authorName}");
                Console.WriteLine($"Created: {created}");
                Console.WriteLine(new string('-', 40));
            }

            // Save the workbook (optional, just to demonstrate lifecycle usage)
            workbook.Save("ThreadedCommentsDemo.xlsx");
        }
    }
}