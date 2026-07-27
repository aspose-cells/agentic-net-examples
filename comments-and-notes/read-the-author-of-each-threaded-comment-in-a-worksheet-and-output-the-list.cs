// Title: Aspose.Cells C# Example – List Threaded Comment Authors in an Excel Worksheet
// Description: Demonstrates how to create a workbook, add threaded comment authors, attach threaded comments to a cell, iterate through all comments, retrieve each threaded comment's author name, output the cell reference and author, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | threaded comment author | Excel worksheet | list comment authors | retrieve comment author | ThreadedCommentAuthorCollection | Aspose.Cells API | Excel collaboration
// Common Searches: Aspose.Cells get threaded comment author C# | list authors of Excel threaded comments using .NET | iterate worksheet comments Aspose.Cells | retrieve comment author names Aspose.Cells | C# example for threaded comment authors in Excel
// Developer Intent: Extract and display the author name of every threaded comment in a worksheet.
// Use Cases: Create an audit log of user contributions in shared Excel files | Build a UI that shows participants of each comment thread | Export comment author data for compliance reporting | Analyze collaboration patterns across workbook cells
// AI Prompts: Provide C# code to filter threaded comments by a specific author with Aspose.Cells. | Show how to export threaded comment authors and their messages to CSV using Aspose.Cells. | Explain how to add a new threaded comment author and retrieve its index in Aspose.Cells for .NET. | Generate a summary of comment threads grouped by author using Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, add threaded comment authors, attach threaded comments to a cell, iterate through all comments, retrieve each threaded comment's author name, output the cell reference and author, and save the file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add some threaded comment authors
            ThreadedCommentAuthorCollection authors = workbook.Worksheets.ThreadedCommentAuthors;
            int aliceIdx = authors.Add("Alice", "alice@example.com", "A");
            int bobIdx = authors.Add("Bob", "bob@example.com", "B");
            ThreadedCommentAuthor alice = authors[aliceIdx];
            ThreadedCommentAuthor bob = authors[bobIdx];

            // Add a regular comment to cell A1 (required to hold threaded comments)
            int commentIdx = worksheet.Comments.Add("A1");
            Comment comment = worksheet.Comments[commentIdx];
            comment.Note = "Parent comment";

            // Add threaded comments with different authors
            comment.ThreadedComments.Add("First threaded comment", alice);
            comment.ThreadedComments.Add("Second threaded comment", bob);

            // Iterate over all comments in the worksheet
            foreach (Comment c in worksheet.Comments)
            {
                // Get the collection of threaded comments for the current comment
                ThreadedCommentCollection threadedComments = c.ThreadedComments;

                // Output the author of each threaded comment
                for (int i = 0; i < threadedComments.Count; i++)
                {
                    ThreadedComment tc = threadedComments[i];
                    // Convert row/column to cell name (e.g., "A1")
                    string cellName = CellsHelper.CellIndexToName(c.Row, c.Column);
                    Console.WriteLine($"Cell {cellName} - Threaded comment {i + 1} author: {tc.Author.Name}");
                }
            }

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("ThreadedCommentsAuthorsDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
