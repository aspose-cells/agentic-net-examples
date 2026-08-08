// Title: Add a Multi‑Line Threaded Comment to Cell H2 with Aspose.Cells for .NET
// Description: Shows how to create a Workbook, add a threaded comment author, define a string with \r\n line breaks, place a multi‑line threaded comment into cell H2 (row 2, column 8), and save the file as ThreadedComment_H2.xlsx using Aspose.Cells in C#.
// Keywords: Aspose.Cells | C# threaded comment | Excel multi‑line comment | preserve line breaks | cell H2 | add comment author | Aspose.Cells .NET example | threaded comment API | Excel automation | programmatic Excel comments
// Common Searches: Aspose.Cells add threaded comment C# | multi line comment Excel .NET | preserve newline in Aspose.Cells comment | threaded comment author Aspose.Cells | how to comment cell H2 programmatically | Excel comment line break Aspose | C# example threaded comment
// Developer Intent: Insert a threaded comment containing several lines into cell H2, keep the line breaks, and assign a custom author.
// Use Cases: Add detailed review notes to a specific cell for collaborative report editing. | Programmatically annotate generated data with multi‑line explanations for auditors. | Highlight data anomalies with author‑attributed threaded comments in financial models.
// AI Prompts: Write C# code using Aspose.Cells to add a threaded comment with newline characters to cell H2 and set a custom author. | Explain how Aspose.Cells preserves line breaks in threaded comments and show alternative formatting options for multi‑line text. | Provide a step‑by‑step guide to create, edit, and delete threaded comments that contain multi‑line content in an Excel workbook with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentExample
{
    // Shows how to create a Workbook, add a threaded comment author, define a string with \r\n line breaks, place a multi‑line threaded comment into cell H2 (row 2, column 8), and save the file as ThreadedComment_H2.xlsx using Aspose.Cells in C#.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a threaded comment author
            // Parameters: Name, UserId, ProviderId
            int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add("Demo Author", "demo_user", "demo_provider");
            ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

            // Define multi-line comment text (preserve line breaks)
            string multiLineText = "First line of comment.\r\nSecond line of comment.\r\nThird line of comment.";

            // Add a threaded comment to cell H2 (row index 1, column index 7)
            worksheet.Comments.AddThreadedComment(1, 7, multiLineText, author);

            // Save the workbook
            workbook.Save("ThreadedComment_H2.xlsx");
        }
    }
}
