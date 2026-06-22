using System;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Merge the range A1:B2 (rows 0-1, columns 0-1)
            // totalRows and totalColumns are 2 (one‑based count)
            worksheet.Cells.Merge(0, 0, 2, 2);

            // Add a threaded comment author
            int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add(
                "Demo Author",          // Author name
                "demoUser@example.com",// User ID / email
                "DemoProvider");       // Provider ID
            ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

            // Add a threaded comment to the merged cell (upper‑left cell A1)
            worksheet.Comments.AddThreadedComment(0, 0,
                "This is a centered threaded comment.", author);

            // Retrieve the comment object to set alignment and visibility
            Comment comment = worksheet.Comments[0, 0];
            comment.TextHorizontalAlignment = TextAlignmentType.Center;
            comment.TextVerticalAlignment   = TextAlignmentType.Center;
            comment.IsVisible = true; // Ensure the comment is shown

            // Save the workbook
            workbook.Save("MergedCellThreadedComment.xlsx");
        }
    }
}