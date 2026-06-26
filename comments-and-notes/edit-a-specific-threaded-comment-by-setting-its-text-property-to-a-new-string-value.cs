using System;
using Aspose.Cells;

class EditThreadedComment
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a threaded comment author
        int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add("User1", "user1@example.com", "PROVIDER");
        ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

        // Add a threaded comment to cell A1 (row 0, column 0) with initial text
        worksheet.Comments.AddThreadedComment(0, 0, "Original comment text", author);

        // Retrieve the threaded comment collection for cell A1
        ThreadedCommentCollection threadedComments = worksheet.Comments.GetThreadedComments(0, 0);

        // Assume we want to edit the first threaded comment in the collection
        if (threadedComments.Count > 0)
        {
            ThreadedComment commentToEdit = threadedComments[0];

            // Update the comment's text using the Notes property
            commentToEdit.Notes = "Updated comment text";
        }

        // Save the workbook with the edited threaded comment
        workbook.Save("EditedThreadedComment.xlsx");
    }
}