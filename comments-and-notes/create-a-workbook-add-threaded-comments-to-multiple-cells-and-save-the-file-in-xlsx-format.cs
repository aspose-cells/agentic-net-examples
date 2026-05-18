using System;
using Aspose.Cells;

namespace ThreadedCommentDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a threaded comment author
            ThreadedCommentAuthorCollection authorCollection = workbook.Worksheets.ThreadedCommentAuthors;
            int authorIndex = authorCollection.Add("John Doe", "john.doe@example.com", "PROVIDER");
            ThreadedCommentAuthor author = authorCollection[authorIndex];

            // Access the comments collection of the worksheet
            CommentCollection comments = worksheet.Comments;

            // Add threaded comments to multiple cells
            comments.AddThreadedComment("A1", "First threaded comment", author);
            comments.AddThreadedComment("B2", "Second threaded comment", author);
            comments.AddThreadedComment("C3", "Third threaded comment", author);

            // Save the workbook in XLSX format
            workbook.Save("ThreadedCommentsDemo.xlsx");
        }
    }
}