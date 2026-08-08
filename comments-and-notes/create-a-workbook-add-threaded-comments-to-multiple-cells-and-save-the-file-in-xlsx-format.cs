// Title: Add Threaded Comments from Multiple Authors to Cells and Save as XLSX with Aspose.Cells (C#)
// Description: Creates a new workbook, defines two comment authors, inserts threaded comment threads into cells A1, B2 and C3, optionally reads the thread for A1, and saves the file as ThreadedCommentsDemo.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | threaded comments | add comment authors | comment collection | save workbook as xlsx | Excel API | collaborative spreadsheet | programmatic comments | Excel automation
// Common Searches: Aspose.Cells add threaded comment C# | How to create comment authors in Aspose.Cells | Save workbook with comments to XLSX using .NET | Retrieve threaded comments from a cell Aspose.Cells | Example of multiple authors threaded comments Excel
// Developer Intent: Insert threaded comment threads from several authors into specific cells and export the workbook as an XLSX file.
// Use Cases: Prepare a review‑ready report where reviewers leave threaded feedback on key cells before distribution. | Build a collaborative template that includes pre‑populated comment threads for onboarding or training purposes. | Programmatically audit all comments on a cell (e.g., A1) to generate a feedback summary before saving.
// AI Prompts: Generate C# code with Aspose.Cells that adds three authors, creates a threaded comment thread on cell D4, and saves the workbook as XLSX. | Explain how to extract all threaded comments from a range of cells and output them as JSON using Aspose.Cells. | Provide step‑by‑step instructions to modify the text of an existing threaded comment in a workbook with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Creates a new workbook, defines two comment authors, inserts threaded comment threads into cells A1, B2 and C3, optionally reads the thread for A1, and saves the file as ThreadedCommentsDemo.xlsx using Aspose.Cells for .NET.
class ThreadedCommentsDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add authors to the workbook
        ThreadedCommentAuthorCollection authors = workbook.Worksheets.ThreadedCommentAuthors;
        int author1Idx = authors.Add("Alice", "alice@example.com", "PROV1");
        int author2Idx = authors.Add("Bob", "bob@example.com", "PROV2");
        ThreadedCommentAuthor author1 = authors[author1Idx];
        ThreadedCommentAuthor author2 = authors[author2Idx];

        // Access the comments collection of the worksheet
        CommentCollection comments = worksheet.Comments;

        // Add threaded comments to multiple cells using the cell-name overload
        comments.AddThreadedComment("A1", "First comment by Alice", author1);
        comments.AddThreadedComment("A1", "Reply by Bob", author2);
        comments.AddThreadedComment("B2", "Bob's comment on B2", author2);
        comments.AddThreadedComment("C3", "Alice adds note to C3", author1);

        // Retrieve and display threaded comments for cell A1 (optional)
        ThreadedCommentCollection a1Threaded = comments.GetThreadedComments("A1");
        foreach (ThreadedComment tc in a1Threaded)
        {
            Console.WriteLine($"Cell A1 - {tc.Author.Name}: {tc.Notes}");
        }

        // Save the workbook in XLSX format
        workbook.Save("ThreadedCommentsDemo.xlsx");
    }
}
