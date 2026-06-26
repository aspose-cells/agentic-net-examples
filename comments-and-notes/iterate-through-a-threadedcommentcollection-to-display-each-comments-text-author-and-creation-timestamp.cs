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

            // Add a threaded comment author to the workbook
            int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add(
                "John Doe",               // Author name
                "john.doe@example.com",   // User ID / email
                "Provider1");             // Provider ID
            ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

            // Add a regular comment to cell A1 (required to hold threaded comments)
            int commentIndex = worksheet.Comments.Add("A1");
            Comment comment = worksheet.Comments[commentIndex];

            // Add several threaded comments to the comment
            comment.ThreadedComments.Add("First threaded comment", author);
            comment.ThreadedComments.Add("Second threaded comment", author);
            comment.ThreadedComments.Add("Third threaded comment", author);

            // Optionally set creation times (if not set, Aspose.Cells assigns the current time)
            foreach (ThreadedComment tc in comment.ThreadedComments)
            {
                tc.CreatedTime = DateTime.Now;
            }

            // Retrieve the ThreadedCommentCollection
            ThreadedCommentCollection threadedComments = comment.ThreadedComments;

            // Iterate through the collection and display text, author name, and creation timestamp
            foreach (ThreadedComment tc in threadedComments)
            {
                string text = tc.Notes;
                string authorName = tc.Author != null ? tc.Author.Name : "Unknown";
                DateTime created = tc.CreatedTime;

                Console.WriteLine($"Text: {text}");
                Console.WriteLine($"Author: {authorName}");
                Console.WriteLine($"Created: {created}");
                Console.WriteLine(new string('-', 40));
            }

            // Save the workbook to verify that comments are persisted
            workbook.Save("ThreadedCommentsOutput.xlsx");
        }
    }
}