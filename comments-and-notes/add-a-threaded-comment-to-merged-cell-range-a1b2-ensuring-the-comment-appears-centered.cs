// Title: Add a Centered Threaded Comment to Merged Cells A1:B2 using Aspose.Cells for .NET
// Description: Shows how to merge cells A1:B2, create a threaded‑comment author, insert a threaded comment, and set both horizontal and vertical alignment to center with Aspose.Cells in C#.
// Keywords: Aspose.Cells | C# threaded comment | merged cells comment | centered comment alignment | ThreadedCommentAuthor | AddThreadedComment | Excel comment API | Aspose.Cells .NET | comment alignment | merge cells Aspose
// Common Searches: Aspose.Cells add threaded comment to merged cell | center comment in merged cells C# | set threaded comment alignment Aspose.Cells | how to create threaded comment author .NET | merge A1:B2 and add comment Aspose.Cells
// Developer Intent: Insert a threaded comment into a merged range and align its text to the center.
// Use Cases: Generate collaborative Excel reports where merged header cells display centered threaded comments. | Automate workbook creation with merged titles and attached review notes. | Programmatically adjust comment positioning after merging cells to maintain visual consistency.
// AI Prompts: Write C# code with Aspose.Cells that merges A1:B2 and adds a centered threaded comment. | Show how to create a ThreadedCommentAuthor and set TextHorizontalAlignment and TextVerticalAlignment for a comment on a merged range. | Explain the steps to retrieve a comment object after adding it to a merged cell and center its text.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsThreadedCommentExample
{
    // Shows how to merge cells A1:B2, create a threaded‑comment author, insert a threaded comment, and set both horizontal and vertical alignment to center with Aspose.Cells in C#.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Merge the range A1:B2 (rows 0-1, columns 0-1)
            worksheet.Cells.Merge(0, 0, 2, 2);

            // Add a threaded comment author
            int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add("Demo Author", "demoUser", "provider");
            ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

            // Add a threaded comment to the merged cell (use the upper‑left cell address)
            worksheet.Comments.AddThreadedComment("A1", "This is a centered threaded comment.", author);

            // Retrieve the comment object to set alignment
            Comment comment = worksheet.Comments["A1"];
            comment.TextHorizontalAlignment = TextAlignmentType.Center;
            comment.TextVerticalAlignment = TextAlignmentType.Center;

            // Save the workbook
            workbook.Save("MergedCellThreadedComment.xlsx");
        }
    }
}
