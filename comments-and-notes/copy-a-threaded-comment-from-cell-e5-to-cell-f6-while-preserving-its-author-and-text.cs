using System;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentCopy
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // Prepare a sample threaded comment in cell E5 (row 4, column 4)
            // ------------------------------------------------------------
            // Add an author to the workbook (if not already present)
            int authorIdx = worksheet.Workbook.Worksheets.ThreadedCommentAuthors.Add(
                "John Doe",               // Name
                "john.doe@example.com",  // UserId
                "EXAMPLE_PROVIDER");     // ProviderId
            ThreadedCommentAuthor author = worksheet.Workbook.Worksheets.ThreadedCommentAuthors[authorIdx];

            // Add a threaded comment to cell E5
            worksheet.Comments.AddThreadedComment("E5", "Original threaded comment text", author);

            // ------------------------------------------------------------
            // Copy all threaded comments from E5 to F6 while preserving author and text
            // ------------------------------------------------------------
            CommentCollection comments = worksheet.Comments;

            // Retrieve the collection of threaded comments from the source cell (E5)
            ThreadedCommentCollection sourceThreadedComments = comments.GetThreadedComments("E5");

            // Iterate through each threaded comment and add it to the destination cell (F6)
            foreach (ThreadedComment srcComment in sourceThreadedComments)
            {
                // Preserve the original author and comment text
                comments.AddThreadedComment("F6", srcComment.Notes, srcComment.Author);
            }

            // ------------------------------------------------------------
            // Save the workbook to verify the result
            // ------------------------------------------------------------
            workbook.Save("CopyThreadedComment.xlsx");
        }
    }
}