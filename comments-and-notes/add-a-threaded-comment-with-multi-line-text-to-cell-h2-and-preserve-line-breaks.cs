using System;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a threaded comment author
            int authorIndex = worksheet.Workbook.Worksheets.ThreadedCommentAuthors.Add(
                "Demo Author",          // Author name
                "demo.author@example.com", // User ID (email)
                "DEMO_PROVIDER");      // Provider ID
            ThreadedCommentAuthor author = worksheet.Workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

            // Multi-line comment text (preserves line breaks)
            string commentText = "First line of comment.\nSecond line of comment.\nThird line of comment.";

            // Add a threaded comment to cell H2 (row index 1, column index 7)
            worksheet.Comments.AddThreadedComment(1, 7, commentText, author);

            // Save the workbook
            workbook.Save("ThreadedComment_H2.xlsx");
        }
    }
}