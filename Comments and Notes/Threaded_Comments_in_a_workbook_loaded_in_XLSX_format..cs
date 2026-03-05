using System;
using Aspose.Cells;

namespace ThreadedCommentsDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook
            Workbook workbook = new Workbook("Input.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];

            // Obtain the collection of threaded comment authors
            ThreadedCommentAuthorCollection authors = workbook.Worksheets.ThreadedCommentAuthors;

            // Add a new author (or reuse an existing one)
            int authorIndex = authors.Add("John Doe", "john.doe@example.com", "JD");
            ThreadedCommentAuthor author = authors[authorIndex];

            // Add a threaded comment to cell B2 (row 1, column 1)
            worksheet.Comments.AddThreadedComment(1, 1, "Initial threaded comment", author);

            // Add a reply to the same cell using the same author for simplicity
            worksheet.Comments.AddThreadedComment(1, 1, "Reply to the initial comment", author);

            // Retrieve all threaded comments for cell B2 by row/column
            ThreadedCommentCollection threadedComments = worksheet.Comments.GetThreadedComments(1, 1);

            // Display each threaded comment's text and author name
            Console.WriteLine("Threaded comments in cell B2:");
            foreach (ThreadedComment tc in threadedComments)
            {
                Console.WriteLine($"- {tc.Notes} (by {tc.Author.Name})");
            }

            // Save the workbook with the new threaded comments
            workbook.Save("Output_WithThreadedComments.xlsx");
        }
    }
}