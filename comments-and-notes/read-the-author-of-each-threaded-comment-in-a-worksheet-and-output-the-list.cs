using System;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentAuthors
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // Prepare sample data: add two authors and some threaded comments
            // ------------------------------------------------------------
            ThreadedCommentAuthorCollection authors = workbook.Worksheets.ThreadedCommentAuthors;

            // Add authors to the collection
            int aliceIndex = authors.Add("Alice", "alice@example.com", "A");
            int bobIndex   = authors.Add("Bob",   "bob@example.com",   "B");

            // Retrieve author objects
            ThreadedCommentAuthor alice = authors[aliceIndex];
            ThreadedCommentAuthor bob   = authors[bobIndex];

            // Add a regular comment to cell A1
            Comment comment = worksheet.Comments[worksheet.Comments.Add("A1")];
            comment.Note = "Parent comment";

            // Add threaded comments to the comment
            comment.ThreadedComments.Add("First threaded comment", alice);
            comment.ThreadedComments.Add("Second threaded comment", bob);

            // ------------------------------------------------------------
            // Read and output the author of each threaded comment in the worksheet
            // ------------------------------------------------------------
            foreach (Comment cmt in worksheet.Comments)
            {
                foreach (ThreadedComment tc in cmt.ThreadedComments)
                {
                    // Output author name
                    Console.WriteLine($"Threaded comment author: {tc.Author.Name}");
                }
            }

            // Save the workbook (lifecycle rule)
            workbook.Save("ThreadedCommentAuthorsOutput.xlsx");
        }
    }
}