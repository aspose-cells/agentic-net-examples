// Title: C# – Iterate ThreadedCommentCollection to display text, author and timestamp with Aspose.Cells
// Description: Creates a workbook, adds a threaded comment author, attaches three threaded comments to cell A1, retrieves the ThreadedCommentCollection, and loops through it to print each comment's text, author name, and creation time before saving the file.
// Keywords: Aspose.Cells | ThreadedCommentCollection | C# | iterate threaded comments | comment author | creation timestamp | worksheet comments | loop through comments
// Common Searches: how to loop through threaded comments Aspose.Cells C# | retrieve author name from ThreadedCommentCollection | display created time of threaded comment Aspose.Cells | list all threaded comments in a cell
// Developer Intent: Enumerate every threaded comment attached to a cell and output its text, author, and creation date.
// Use Cases: Generate a console audit of discussion threads for a specific cell. | Export comment details (text, author, date) to a log or database for change tracking. | Validate that each threaded comment contains author information before publishing the workbook.
// AI Prompts: Write C# code to filter a ThreadedCommentCollection by author using Aspose.Cells. | Show how to sort threaded comments by CreatedTime and print them chronologically. | Provide an example that exports threaded comment text, author, and timestamp to a CSV file with Aspose.Cells.

using System;
using Aspose.Cells;

namespace ThreadedCommentIterationDemo
{
    // Creates a workbook, adds a threaded comment author, attaches three threaded comments to cell A1, retrieves the ThreadedCommentCollection, and loops through it to print each comment's text, author name, and creation time before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a threaded comment author (author collection rule)
            int authorIdx = workbook.Worksheets.ThreadedCommentAuthors.Add("Alice", "alice@example.com", "Provider1");
            ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIdx];

            // Add a regular comment to cell A1 (comment collection rule)
            int commentIdx = worksheet.Comments.Add("A1");
            Comment comment = worksheet.Comments[commentIdx];

            // Add several threaded comments to the comment (ThreadedCommentCollection.Add rule)
            comment.ThreadedComments.Add("First threaded comment", author);
            comment.ThreadedComments.Add("Second threaded comment", author);
            comment.ThreadedComments.Add("Third threaded comment", author);

            // Retrieve the threaded comments collection
            ThreadedCommentCollection threadedComments = comment.ThreadedComments;

            // Iterate through the collection and display text, author name, and creation time
            foreach (ThreadedComment tc in threadedComments)
            {
                Console.WriteLine($"Text: {tc.Notes}");
                Console.WriteLine($"Author: {tc.Author.Name}");
                Console.WriteLine($"Created: {tc.CreatedTime}");
                Console.WriteLine(new string('-', 40));
            }

            // Save the workbook (save rule)
            workbook.Save("ThreadedCommentsDemo.xlsx");
        }
    }
}
