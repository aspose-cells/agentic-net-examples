using System;
using Aspose.Cells;

class RemoveThreadedCommentDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a threaded comment author
        int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add(
            "John Doe",               // author name
            "john.doe@example.com",   // author email / user id
            "PROVIDER");              // provider id
        ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

        // Add a threaded comment to cell C3
        worksheet.Comments.AddThreadedComment("C3", "Initial threaded comment", author);

        // Remove the comment (and its threaded comments) from cell C3
        worksheet.Comments.RemoveAt("C3");

        // Save the workbook
        workbook.Save("RemovedThreadedComment.xlsx");
    }
}