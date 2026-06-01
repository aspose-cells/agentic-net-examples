using System;
using Aspose.Cells;

class ThreadedCommentIterationDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a threaded comment author to the workbook
        int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add(
            "Alice",               // Author name
            "alice@example.com",   // User ID / email
            "Provider1");          // Provider ID
        ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

        // Add a regular comment to cell A1 (required to host threaded comments)
        int commentIndex = worksheet.Comments.Add("A1");
        Comment comment = worksheet.Comments[commentIndex];

        // Add threaded comments to the comment
        comment.ThreadedComments.Add("First threaded comment", author);
        comment.ThreadedComments.Add("Second threaded comment", author);

        // Retrieve the collection of threaded comments
        ThreadedCommentCollection threadedComments = comment.ThreadedComments;

        // Iterate through the collection and display each comment's details
        for (int i = 0; i < threadedComments.Count; i++)
        {
            ThreadedComment tc = threadedComments[i];
            Console.WriteLine($"Comment {i + 1}:");
            Console.WriteLine($"  Text   : {tc.Notes}");
            Console.WriteLine($"  Author : {tc.Author.Name}");
            Console.WriteLine($"  Created: {tc.CreatedTime}");
        }

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("ThreadedCommentsDemo.xlsx");
    }
}