// Title: Add a Centered Threaded Comment to a Merged Cell (A1:B2) Using Aspose.Cells for .NET
// Description: This example merges cells A1:B2, creates a threaded‑comment author, inserts a threaded comment on the merged area, aligns the text horizontally and vertically, makes the comment visible, and saves the workbook as ThreadedCommentMergedCell.xlsx with Aspose.Cells in C#.
// Keywords: Aspose.Cells | C# | threaded comment | merged cells | centered comment | comment alignment | Excel comment API | Aspose.Cells ThreadedComment | programmatic Excel comment | merge cells with comment
// Common Searches: how to add a threaded comment to merged cells using Aspose.Cells | center comment text in merged range Aspose.Cells C# | set threaded comment visibility and alignment Aspose.Cells | Aspose.Cells example for merged cell comments | C# code to create threaded comment author in Aspose.Cells
// Developer Intent: Insert a threaded comment into a merged cell range and center its text.
// Use Cases: Add explanatory notes to a merged header row in an automated report. | Place reviewer feedback directly on a merged title block of a spreadsheet template. | Pre‑populate a workbook with centered threaded comments for data validation sections.
// AI Prompts: Generate C# code with Aspose.Cells that merges A1:C3 and adds a centered threaded comment. | Show how to make a threaded comment visible and align its text after merging cells in Aspose.Cells. | Explain the steps to create a ThreadedCommentAuthor and assign it to a comment on a merged cell range.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsThreadedCommentDemo
{
    // This example merges cells A1:B2, creates a threaded‑comment author, inserts a threaded comment on the merged area, aligns the text horizontally and vertically, makes the comment visible, and saves the workbook as ThreadedCommentMergedCell.xlsx with Aspose.Cells in C#.
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
            int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add(
                "Demo Author",          // Author name
                "demoUser",             // User ID
                "demoProvider");        // Provider
            ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

            // Add a threaded comment to the merged cell (use the upper‑left cell address "A1")
            int commentIndex = worksheet.Comments.AddThreadedComment(
                "A1",                                   // Cell address
                "This is a centered threaded comment.", // Comment text
                author);                                // Author

            // Retrieve the comment object to set its appearance
            Comment comment = worksheet.Comments[commentIndex];
            comment.TextHorizontalAlignment = TextAlignmentType.Center;
            comment.TextVerticalAlignment = TextAlignmentType.Center;
            comment.IsVisible = true; // Make the comment visible by default

            // Save the workbook
            workbook.Save("ThreadedCommentMergedCell.xlsx");
        }
    }
}
