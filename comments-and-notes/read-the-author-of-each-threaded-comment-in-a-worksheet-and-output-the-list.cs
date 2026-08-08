// Title: Aspose.Cells C# – List All Threaded Comment Authors in a Worksheet
// Description: Demonstrates how to create a workbook, add threaded comment authors, insert threaded comments, and iterate through every comment to output each author’s name with its cell address, then save the file.
// Keywords: Aspose.Cells | C# | .NET | threaded comment authors | ThreadedCommentAuthorCollection | read comment author | list comment authors | iterate worksheet comments | Excel audit | global developers
// Common Searches: how to get threaded comment author Aspose.Cells C# | list all comment authors in Excel using Aspose.Cells | iterate worksheet comments .NET example | Aspose.Cells retrieve threaded comment author name | C# code to enumerate Excel comment authors
// Developer Intent: Extract and display the author name for every threaded comment in a worksheet.
// Use Cases: Create an audit report of who commented on each cell. | Validate that only permitted users have added threaded comments before publishing. | Export cell coordinates with corresponding comment authors for compliance tracking.
// AI Prompts: Generate C# code that collects all threaded comment authors from a worksheet and writes the author‑cell pairs to a CSV file using Aspose.Cells. | Show an example that filters threaded comments by a specific author name and returns the affected cell addresses. | Explain how to replace a threaded comment author with another author across an entire workbook in C#.

using System;
using Aspose.Cells;

namespace ThreadedCommentAuthorsDemo
{
    // Demonstrates how to create a workbook, add threaded comment authors, insert threaded comments, and iterate through every comment to output each author’s name with its cell address, then save the file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the collection of threaded comment authors
            ThreadedCommentAuthorCollection authors = workbook.Worksheets.ThreadedCommentAuthors;

            // Add two authors to the collection
            int author1Index = authors.Add("Alice", "alice@example.com", "A");
            int author2Index = authors.Add("Bob", "bob@example.com", "B");
            ThreadedCommentAuthor author1 = authors[author1Index];
            ThreadedCommentAuthor author2 = authors[author2Index];

            // Add threaded comments to some cells using the authors
            worksheet.Comments.AddThreadedComment("A1", "First comment", author1);
            worksheet.Comments.AddThreadedComment("A1", "Reply comment", author2);
            worksheet.Comments.AddThreadedComment("B2", "Another comment", author2);

            // Iterate through all comments in the worksheet
            CommentCollection commentCollection = worksheet.Comments;
            for (int i = 0; i < commentCollection.Count; i++)
            {
                // Each comment may have multiple threaded comments
                Comment comment = commentCollection[i];
                ThreadedCommentCollection threadedComments = comment.ThreadedComments;

                for (int j = 0; j < threadedComments.Count; j++)
                {
                    ThreadedComment threadedComment = threadedComments[j];
                    // Output the author name of each threaded comment
                    Console.WriteLine($"Cell ({comment.Row}, {comment.Column}) - Threaded Comment {j + 1} Author: {threadedComment.Author.Name}");
                }
            }

            // Save the workbook (optional, just to demonstrate lifecycle usage)
            workbook.Save("ThreadedCommentsAuthorsOutput.xlsx");
        }
    }
}
