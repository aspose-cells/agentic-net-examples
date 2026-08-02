// Title: Add Threaded Comment to Cell J12 and Reply Programmatically with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a new workbook, define two threaded‑comment authors, insert an initial comment in cell J12, retrieve the comment thread, append a reply from a second author, and save the file as ThreadedComment_J12.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells threaded comment C# | add comment to cell J12 | reply to threaded comment .NET | ThreadedCommentAuthors Aspose | save workbook with comments | programmatic comment thread | Excel collaboration Aspose
// Common Searches: how to add a threaded comment to a specific cell using Aspose.Cells C# | Aspose.Cells reply to threaded comment programmatically | define multiple authors for threaded comments in Aspose.Cells | retrieve and reply to comments in Excel cell J12 with Aspose | C# example for threaded comments in Aspose.Cells
// Developer Intent: Create a threaded comment on cell J12, then add a reply from another author, all via Aspose.Cells for .NET.
// Use Cases: Capture review notes by inserting an initial comment from Alice in cell J12. | Document discussion by adding Bob’s reply to the same comment thread. | Share the workbook with a complete comment conversation for collaborative editing.
// AI Prompts: Generate C# code that adds a threaded comment to cell J12 with author Alice and then adds a reply from author Bob using Aspose.Cells. | Explain how to retrieve the ThreadedCommentCollection for a cell and append a reply in Aspose.Cells for .NET. | Show how to create multiple threaded comment authors and assign them to comments in an Aspose.Cells workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentDemo
{
    // Demonstrates how to create a new workbook, define two threaded‑comment authors, insert an initial comment in cell J12, retrieve the comment thread, append a reply from a second author, and save the file as ThreadedComment_J12.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the comments collection of the worksheet
            CommentCollection comments = worksheet.Comments;

            // Add authors for the threaded comments
            int author1Index = workbook.Worksheets.ThreadedCommentAuthors.Add("Alice", "alice@example.com", "PROVIDER_1");
            ThreadedCommentAuthor author1 = workbook.Worksheets.ThreadedCommentAuthors[author1Index];

            int author2Index = workbook.Worksheets.ThreadedCommentAuthors.Add("Bob", "bob@example.com", "PROVIDER_2");
            ThreadedCommentAuthor author2 = workbook.Worksheets.ThreadedCommentAuthors[author2Index];

            // Add a threaded comment to cell J12 (row 11, column 9 – zero‑based indexing)
            comments.AddThreadedComment(11, 9, "Initial comment on J12.", author1);

            // Retrieve the threaded comment collection for J12
            ThreadedCommentCollection threadedComments = comments.GetThreadedComments(11, 9);

            // Add a reply (additional note) to the same threaded comment thread
            threadedComments.Add("This is a reply from Bob.", author2);

            // Save the workbook
            workbook.Save("ThreadedComment_J12.xlsx");
        }
    }
}
