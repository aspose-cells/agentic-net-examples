using System;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentCreatedTimeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a threaded comment author
            int authorId = workbook.Worksheets.ThreadedCommentAuthors.Add("User1", "user1@example.com", "PROVIDER1");
            ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorId];

            // Add a threaded comment to cell A1
            worksheet.Comments.AddThreadedComment("A1", "Sample threaded comment", author);

            // Retrieve the threaded comment we just added
            ThreadedComment threadedComment = worksheet.Comments[0].ThreadedComments[0];

            // Get the creation time of the threaded comment using the CreatedTime property
            DateTime createdTime = threadedComment.CreatedTime;

            // Log the creation time to the console
            Console.WriteLine("Threaded comment created at: " + createdTime);

            // Save the workbook (lifecycle: save)
            workbook.Save("ThreadedCommentCreatedTimeDemo.xlsx");
        }
    }
}