using System;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentCount
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a threaded comment author
            int authorIdx = workbook.Worksheets.ThreadedCommentAuthors.Add("John Doe", "john.doe@example.com", "JD");
            ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIdx];

            // Add some threaded comments to different cells
            sheet.Comments.AddThreadedComment("A1", "First comment", author);
            sheet.Comments.AddThreadedComment("A1", "Reply to first", author);
            sheet.Comments.AddThreadedComment("B2", "Another comment", author);

            // Variable to hold total count
            int totalThreadedComments = 0;

            // Iterate through all worksheets
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Iterate through all comments in the worksheet
                foreach (Comment comment in ws.Comments)
                {
                    // Each comment may contain a collection of threaded comments
                    totalThreadedComments += comment.ThreadedComments.Count;
                }
            }

            // Display the total number of threaded comments
            Console.WriteLine($"Total threaded comments in the workbook: {totalThreadedComments}");

            // Save the workbook (optional, just to demonstrate lifecycle usage)
            workbook.Save("ThreadedCommentsCount.xlsx");
        }
    }
}