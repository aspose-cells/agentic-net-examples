// Title: Update a Threaded Comment’s Text in Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a threaded comment author, insert a threaded comment into cell A1, retrieve the comment collection, modify the comment’s Notes property, and save the file as EditedThreadedComment.xlsx using C# and Aspose.Cells.
// Keywords: Aspose.Cells threaded comment edit | C# update comment notes | modify threaded comment text | Aspose.Cells Notes property | Excel comment editing .NET
// Common Searches: how to change threaded comment text Aspose.Cells | update Notes property of a comment in C# | edit existing threaded comment in Excel with Aspose | Aspose.Cells replace comment content programmatically | C# sample for editing threaded comments
// Developer Intent: Change the text of an existing threaded comment in a worksheet without recreating the comment.
// Use Cases: Fix typographical errors in previously added comments. | Swap placeholder text with final review feedback. | Translate comment content for localized workbook distribution.
// AI Prompts: Generate C# code that retrieves the first threaded comment from cell A1 using Aspose.Cells and updates its Notes property. | Show an example that adds a threaded comment author, creates a comment, edits the comment’s text, and saves the workbook. | Explain best practices for safely updating a specific threaded comment when multiple comments exist in the same cell.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, add a threaded comment author, insert a threaded comment into cell A1, retrieve the comment collection, modify the comment’s Notes property, and save the file as EditedThreadedComment.xlsx using C# and Aspose.Cells.
class EditThreadedComment
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a threaded comment author
        int authorIndex = worksheet.Workbook.Worksheets.ThreadedCommentAuthors.Add("User1", "user1@example.com", "PROVIDER");
        ThreadedCommentAuthor author = worksheet.Workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

        // Add a threaded comment to cell A1 (row 0, column 0) with initial text
        worksheet.Comments.AddThreadedComment(0, 0, "Original comment text", author);

        // Retrieve the threaded comments for cell A1
        ThreadedCommentCollection threadedComments = worksheet.Comments.GetThreadedComments(0, 0);

        // Edit the first threaded comment's text (Notes property)
        if (threadedComments.Count > 0)
        {
            ThreadedComment commentToEdit = threadedComments[0];
            commentToEdit.Notes = "Updated comment text";
        }

        // Save the workbook
        workbook.Save("EditedThreadedComment.xlsx");
    }
}
