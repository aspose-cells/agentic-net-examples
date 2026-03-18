using System;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentUpdate
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a threaded comment author (required for threaded comments)
            int authorIdx = workbook.Worksheets.ThreadedCommentAuthors.Add("Original Author", "authorId", "providerId");
            ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIdx];

            // Add a threaded comment to cell B2 (row 1, column 1)
            int commentIdx = worksheet.Comments.AddThreadedComment(1, 1, "Initial threaded comment text", author);

            // Retrieve the collection of threaded comments for the cell
            ThreadedCommentCollection threadedComments = worksheet.Comments.GetThreadedComments(1, 1);

            // Assuming we want to modify the first (and only) threaded comment
            if (threadedComments.Count > 0)
            {
                ThreadedComment tc = threadedComments[0];
                // Update the comment text
                tc.Notes = "Updated threaded comment text";
                // Optionally, change the author
                int newAuthorIdx = workbook.Worksheets.ThreadedCommentAuthors.Add("Updated Author", "newAuthorId", "providerId");
                ThreadedCommentAuthor newAuthor = workbook.Worksheets.ThreadedCommentAuthors[newAuthorIdx];
                tc.Author = newAuthor;
            }

            // Save the workbook with the modified threaded comment
            workbook.Save("ModifiedThreadedComment.xlsx");
        }
    }
}