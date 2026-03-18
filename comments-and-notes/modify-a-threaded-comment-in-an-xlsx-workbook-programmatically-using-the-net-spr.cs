using System;
using Aspose.Cells;

class ModifyThreadedComment
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("Input.xlsx");

        // Access the first worksheet (or any specific worksheet as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the target cell coordinates (e.g., B2 -> row 1, column 1)
        int targetRow = 1;
        int targetColumn = 1;

        // Retrieve the collection of threaded comments for the specified cell
        ThreadedCommentCollection threadedComments = worksheet.Comments.GetThreadedComments(targetRow, targetColumn);

        // Ensure there is at least one threaded comment to modify
        if (threadedComments != null && threadedComments.Count > 0)
        {
            // Get the first threaded comment (you can choose any index as required)
            ThreadedComment comment = threadedComments[0];

            // Modify the comment text
            comment.Notes = "Updated comment text via Aspose.Cells";

            // Optionally, modify the author name (demonstrates setting the Author property)
            // Retrieve an existing author or create a new one
            ThreadedCommentAuthorCollection authors = workbook.Worksheets.ThreadedCommentAuthors;
            int authorIdx = authors.Add("New Author", "new.author@example.com", "PROVIDER_ID");
            ThreadedCommentAuthor newAuthor = authors[authorIdx];
            comment.Author = newAuthor;

            // Optionally, update the created time
            comment.CreatedTime = DateTime.Now;
        }

        // Save the modified workbook to a new file
        workbook.Save("Output.xlsx");
    }
}