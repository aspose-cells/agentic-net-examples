using System;
using Aspose.Cells;

namespace ThreadedCommentsDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("Input.xlsx");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // -----------------------------------------------------------------
            // 1. Create threaded comment authors (metadata for comments)
            // -----------------------------------------------------------------
            ThreadedCommentAuthorCollection authors = workbook.Worksheets.ThreadedCommentAuthors;

            // Add two authors and keep references
            int authorId1 = authors.Add("Alice Johnson", "alice.johnson@example.com", "PROV1");
            ThreadedCommentAuthor author1 = authors[authorId1];

            int authorId2 = authors.Add("Bob Smith", "bob.smith@example.com", "PROV2");
            ThreadedCommentAuthor author2 = authors[authorId2];

            // -----------------------------------------------------------------
            // 2. Add threaded comments (a thread) to a specific cell, e.g., "C2"
            // -----------------------------------------------------------------
            // First comment – the root of the thread
            worksheet.Comments.AddThreadedComment("C2", "Initial comment from Alice.", author1);

            // Reply comment – second entry in the same thread
            worksheet.Comments.AddThreadedComment("C2", "Reply from Bob.", author2);

            // Another reply from Alice
            worksheet.Comments.AddThreadedComment("C2", "Follow‑up from Alice.", author1);

            // -----------------------------------------------------------------
            // 3. Optionally set additional metadata (e.g., CreatedTime) for each comment
            // -----------------------------------------------------------------
            ThreadedCommentCollection thread = worksheet.Comments.GetThreadedComments("C2");
            for (int i = 0; i < thread.Count; i++)
            {
                ThreadedComment tc = thread[i];
                // Set the creation time to now minus i minutes to simulate different timestamps
                tc.CreatedTime = DateTime.Now.AddMinutes(-i);
            }

            // -----------------------------------------------------------------
            // 4. Display the threaded comments with their metadata
            // -----------------------------------------------------------------
            Console.WriteLine("Threaded comments for cell C2:");
            foreach (ThreadedComment tc in thread)
            {
                Console.WriteLine($"- Author: {tc.Author.Name}");
                Console.WriteLine($"  Notes : {tc.Notes}");
                Console.WriteLine($"  Created: {tc.CreatedTime}");
                Console.WriteLine($"  Location: Row {tc.Row}, Column {tc.Column}");
                Console.WriteLine();
            }

            // -----------------------------------------------------------------
            // 5. Save the workbook with the new threaded comments
            // -----------------------------------------------------------------
            workbook.Save("Output.xlsx");
        }
    }
}