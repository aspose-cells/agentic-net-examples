// Title: C# – Remove Threaded Comment from Cell C3 with Aspose.Cells for .NET
// Description: Shows how to build a workbook, add a threaded‑comment author, place a comment in C3, fetch it via worksheet.Comments["C3"], delete the first entry in its ThreadedComments collection, and save the file without that comment.
// Keywords: Aspose.Cells C# remove threaded comment | delete comment cell C3 Aspose | ThreadedCommentCollection RemoveAt example | Aspose.Cells comment deletion .NET | programmatic spreadsheet comment removal | C# Excel comment API
// Common Searches: how to delete a threaded comment from a specific cell using Aspose.Cells C# | remove first threaded comment in C3 Aspose.Cells | Aspose.Cells .NET remove comment programmatically | C# code to clear threaded comments from Excel cell
// Developer Intent: Programmatically eliminate the threaded comment attached to cell C3 in an Excel workbook using the Aspose.Cells .NET library.
// Use Cases: Strip outdated review notes before publishing a report. | Automate cleanup of temporary feedback in generated spreadsheets. | Prepare a workbook for distribution by removing comments from targeted cells.
// AI Prompts: Demonstrate how to check whether a cell contains threaded comments before attempting removal with Aspose.Cells. | Provide a C# snippet that removes all threaded comments from an entire worksheet. | Explain how to target and delete a specific threaded comment by index when multiple comments exist in one cell.

using System;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentRemoval
{
    // Shows how to build a workbook, add a threaded‑comment author, place a comment in C3, fetch it via worksheet.Comments["C3"], delete the first entry in its ThreadedComments collection, and save the file without that comment.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a threaded comment author
            int authorIdx = workbook.Worksheets.ThreadedCommentAuthors.Add("John Doe", "john.doe@example.com", "EXAMPLE_PROVIDER");
            ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIdx];

            // Add a threaded comment to cell C3 (row 2, column 2)
            worksheet.Comments.AddThreadedComment(2, 2, "Initial threaded comment", author);

            // Retrieve the comment object that resides in C3
            Comment comment = worksheet.Comments["C3"];

            // Access the collection of threaded comments for this comment
            ThreadedCommentCollection threadedComments = comment.ThreadedComments;

            // Remove the first (and only) threaded comment using RemoveAt on the collection
            // This demonstrates removal via the comment object's threaded comment collection
            if (threadedComments.Count > 0)
            {
                threadedComments.RemoveAt(0);
            }

            // Save the workbook to verify that the threaded comment has been removed
            workbook.Save("ThreadedCommentRemoved.xlsx");
        }
    }
}
