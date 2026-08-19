// Title: C# – Iterate ThreadedCommentCollection to show text, author and timestamp with Aspose.Cells
// Description: Creates a workbook, adds a threaded‑comment author, attaches a regular comment to cell A1, inserts several threaded comments, then loops through the ThreadedCommentCollection to output each comment’s text, author name and creation time before saving the file.
// Keywords: Aspose.Cells threaded comments | ThreadedCommentCollection C# | read comment author Aspose | comment creation time Aspose.Cells | iterate threaded comments .NET
// Common Searches: how to loop through ThreadedCommentCollection Aspose.Cells | get author and timestamp from threaded comments C# | display all threaded comments for a cell Aspose | retrieve creation date of threaded comments .NET
// Developer Intent: The developer needs to enumerate a ThreadedCommentCollection and retrieve each comment’s text, author name, and creation timestamp.
// Use Cases: Generate an audit log of every threaded comment in a worksheet with author and date. | Populate a UI grid or list with comment history extracted from a workbook. | Export comment details to CSV/JSON for downstream reporting or analytics.
// AI Prompts: Write C# code using Aspose.Cells to iterate over a ThreadedCommentCollection and print each comment's text, author, and creation time. | Show how to filter threaded comments by a specific author and sort them by CreatedTime in Aspose.Cells for .NET. | Provide an example that exports the text, author, and timestamp of all threaded comments in a worksheet to a CSV file with Aspose.Cells.

using System;
using Aspose.Cells;

namespace ThreadedCommentIterationDemo
{
    // Creates a workbook, adds a threaded‑comment author, attaches a regular comment to cell A1, inserts several threaded comments, then loops through the ThreadedCommentCollection to output each comment’s text, author name and creation time before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a threaded comment author to the workbook
            int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add("John Doe", "john.doe@example.com", "Provider1");
            ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

            // Add a regular comment to cell A1 (required to hold threaded comments)
            int commentIdx = worksheet.Comments.Add("A1");
            Comment comment = worksheet.Comments[commentIdx];

            // Add several threaded comments to the comment
            comment.ThreadedComments.Add("First threaded comment", author);
            comment.ThreadedComments.Add("Second threaded comment", author);
            comment.ThreadedComments.Add("Third threaded comment", author);

            // Retrieve the collection of threaded comments for the cell
            ThreadedCommentCollection threadedComments = comment.ThreadedComments;

            // Iterate through the collection and display text, author name, and creation time
            foreach (ThreadedComment tc in threadedComments)
            {
                string text = tc.Notes;
                string authorName = tc.Author?.Name ?? "Unknown";
                DateTime created = tc.CreatedTime;

                Console.WriteLine($"Text: {text}");
                Console.WriteLine($"Author: {authorName}");
                Console.WriteLine($"Created: {created}");
                Console.WriteLine(new string('-', 40));
            }

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("ThreadedCommentsDemo.xlsx");
        }
    }
}
