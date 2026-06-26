using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a threaded comment author named John
        int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add("John", "john@example.com", "PROVIDER");
        ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

        // Add a threaded comment to cell A1 with the author John
        worksheet.Comments.AddThreadedComment("A1", "This is a threaded comment.", author);

        // Save the workbook
        workbook.Save("ThreadedCommentJohn.xlsx");
    }
}