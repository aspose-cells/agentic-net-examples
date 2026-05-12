using System;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a threaded comment author named John
            // Parameters: name, userId, providerId
            int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add("John", "john@example.com", "PROVIDER");
            ThreadedCommentAuthor johnAuthor = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

            // Add a threaded comment to cell A1 using the author
            // Overload: AddThreadedComment(string cellName, string text, ThreadedCommentAuthor author)
            worksheet.Comments.AddThreadedComment("A1", "This is a threaded comment by John.", johnAuthor);

            // Save the workbook to a file
            workbook.Save("ThreadedCommentDemo.xlsx");
        }
    }
}