using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (creation rule)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add threaded comment authors
        ThreadedCommentAuthorCollection authors = workbook.Worksheets.ThreadedCommentAuthors;
        int aliceIdx = authors.Add("Alice", "alice@example.com", "A");
        int bobIdx = authors.Add("Bob", "bob@example.com", "B");
        ThreadedCommentAuthor alice = authors[aliceIdx];
        ThreadedCommentAuthor bob = authors[bobIdx];

        // Add a base comment to a cell (required to hold threaded comments)
        int baseCommentIdx = worksheet.Comments.Add("A1");
        Comment baseComment = worksheet.Comments[baseCommentIdx];
        baseComment.Note = "Base comment";

        // Add threaded comments with different authors
        baseComment.ThreadedComments.Add("First threaded comment", alice);
        baseComment.ThreadedComments.Add("Second threaded comment", bob);
        baseComment.ThreadedComments.Add("Third threaded comment", alice);

        // Read and output the author of each threaded comment in the worksheet
        Console.WriteLine("Threaded comment authors in the worksheet:");
        foreach (Comment comment in worksheet.Comments)
        {
            ThreadedCommentCollection threadedComments = comment.ThreadedComments;
            for (int i = 0; i < threadedComments.Count; i++)
            {
                ThreadedComment tc = threadedComments[i];
                string authorName = tc.Author != null ? tc.Author.Name : "Unknown";
                Console.WriteLine($"Cell ({comment.Row},{comment.Column}) - Thread {i + 1} Author: {authorName}");
            }
        }

        // Save the workbook (save rule)
        workbook.Save("ThreadedCommentsAuthorsOutput.xlsx");
    }
}