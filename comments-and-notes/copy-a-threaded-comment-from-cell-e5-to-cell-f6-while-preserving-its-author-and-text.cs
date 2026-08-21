// Title: Copy a Threaded Comment from E5 to F6 while retaining author – Aspose.Cells C# Example
// Description: This C# snippet shows how to duplicate a threaded comment from cell E5 to cell F6 in an Aspose.Cells workbook. It creates a sample comment, fetches the ThreadedCommentCollection, adds each comment to the target cell using the original author object, and saves the workbook.
// Keywords: Aspose.Cells threaded comment copy | C# copy comment author | duplicate threaded comment Aspose | preserve comment metadata .NET | AddThreadedComment example | ThreadedCommentCollection usage | Excel comment migration C#
// Common Searches: Aspose.Cells copy threaded comment | how to duplicate Excel comment with author in .NET | move comment from one cell to another Aspose | preserve comment author when copying cells | C# Aspose.Cells threaded comment transfer
// Developer Intent: Duplicate an existing threaded comment from one cell to another without losing the author information.
// Use Cases: Copy review notes when rows are shifted in a financial model | Migrate discussion threads during template redesign | Synchronize comments after data re‑organization across worksheets | Create a backup of comments before bulk data import
// AI Prompts: Write C# code that copies all threaded comments from a source cell to a destination cell in Aspose.Cells, keeping the original author objects. | Explain the role of ThreadedCommentAuthors when cloning comments in Aspose.Cells. | Show how to copy threaded comments and then delete the originals in a single operation using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentCopy
{
    // This C# snippet shows how to duplicate a threaded comment from cell E5 to cell F6 in an Aspose.Cells workbook. It creates a sample comment, fetches the ThreadedCommentCollection, adds each comment to the target cell using the original author object, and saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // -----------------------------------------------------------------
            // Assume a threaded comment already exists in cell E5 (row 4, column 4)
            // For demonstration, we add a sample threaded comment first.
            // -----------------------------------------------------------------
            // Add a threaded comment author
            int authorIdx = worksheet.Workbook.Worksheets.ThreadedCommentAuthors.Add(
                "John Doe",               // Author name
                "john.doe@example.com",  // User ID / email
                "EXAMPLE_PROVIDER");     // Provider ID
            ThreadedCommentAuthor author = worksheet.Workbook.Worksheets.ThreadedCommentAuthors[authorIdx];

            // Add a threaded comment to cell E5
            worksheet.Comments.AddThreadedComment(4, 4, "Original threaded comment text", author);

            // -----------------------------------------------------------------
            // Copy all threaded comments from E5 to F6 while preserving author and text
            // -----------------------------------------------------------------
            int sourceRow = 4;      // E5 row index (zero‑based)
            int sourceColumn = 4;   // E5 column index (zero‑based)

            int destRow = 5;        // F6 row index (zero‑based)
            int destColumn = 5;     // F6 column index (zero‑based)

            // Retrieve the threaded comments collection from the source cell
            ThreadedCommentCollection sourceComments = worksheet.Comments.GetThreadedComments(sourceRow, sourceColumn);

            // Iterate through each threaded comment and add it to the destination cell
            foreach (ThreadedComment srcComment in sourceComments)
            {
                // Preserve the original text and author
                string commentText = srcComment.Notes;
                ThreadedCommentAuthor commentAuthor = srcComment.Author;

                // Add the comment to the destination cell (F6)
                worksheet.Comments.AddThreadedComment(destRow, destColumn, commentText, commentAuthor);
            }

            // Save the workbook to verify the result
            workbook.Save("ThreadedCommentCopyDemo.xlsx");
        }
    }
}
