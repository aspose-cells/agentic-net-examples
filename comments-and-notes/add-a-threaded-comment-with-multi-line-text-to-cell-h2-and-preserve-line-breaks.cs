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

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a threaded comment author
            int authorIndex = worksheet.Workbook.Worksheets.ThreadedCommentAuthors.Add(
                "Demo Author",          // Author name
                "demo.author@example.com", // User ID (email)
                "DEMO_PROVIDER");      // Provider ID
            ThreadedCommentAuthor author = worksheet.Workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

            // Define multi-line comment text (preserves line breaks)
            string multiLineText = "First line of comment.\r\nSecond line of comment.\r\nThird line of comment.";

            // Add a threaded comment to cell H2 (row index 1, column index 7)
            worksheet.Comments.AddThreadedComment(1, 7, multiLineText, author);

            // Save the workbook
            workbook.Save("ThreadedComment_H2.xlsx");
        }
    }
}