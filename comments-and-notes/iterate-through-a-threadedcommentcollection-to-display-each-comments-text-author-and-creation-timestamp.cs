// Title: How to loop through Aspose.Cells ThreadedCommentCollection in C# to print comment text, author, and creation time
// AI Prompts: Write C# code that opens an Aspose.Cells workbook, accesses a cell's comment thread, and prints each threaded comment's Notes, Author.Name, and CreatedTime. | Provide a full example that adds a threaded comment author, creates a regular comment with multiple threaded replies, then iterates the ThreadedCommentCollection to output comment details using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# retrieve threaded comment author name and timestamp | C# iterate over ThreadedCommentCollection in Aspose.Cells workbook | How to display notes and created time of Excel threaded comments using Aspose.Cells | Sample code for enumerating threaded comments in a worksheet with Aspose.Cells .NET
// Tags: Aspose.Cells iterate ThreadedCommentCollection C# | read threaded comment author Aspose.Cells | extract created time from Aspose.Cells threaded comments | C# enumerate Excel threaded comments using Aspose.Cells | display comment notes Aspose.Cells workbook

using System;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentDemo
{
    // The program creates a workbook, adds a threaded comment author, inserts a regular comment in cell A1, adds three threaded replies, iterates the ThreadedCommentCollection printing each comment's text, author name, and creation timestamp, and finally saves the workbook.
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

            // Retrieve the ThreadedCommentCollection
            ThreadedCommentCollection threadedComments = comment.ThreadedComments;

            // Iterate through the collection and display details
            foreach (ThreadedComment tc in threadedComments)
            {
                // Text of the comment
                string text = tc.Notes;

                // Author name
                string authorName = tc.Author?.Name ?? "Unknown";

                // Creation timestamp (if not set, it defaults to DateTime.MinValue)
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
