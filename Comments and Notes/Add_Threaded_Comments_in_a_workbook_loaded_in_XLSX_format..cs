using System;
using Aspose.Cells;

class ThreadedCommentDemo
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("Input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Create a threaded comment author (name, userId, providerId)
        int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add("John Doe", "jdoe", "PROVIDER");
        ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

        // Add a threaded comment to cell B2 using row/column indices (row 1, column 1)
        worksheet.Comments.AddThreadedComment(1, 1, "This is a threaded comment on B2.", author);

        // Add another threaded comment to cell C3 using the cell name overload
        worksheet.Comments.AddThreadedComment("C3", "Another threaded comment on C3.", author);

        // Retrieve and display all threaded comments for cell B2
        ThreadedCommentCollection b2Comments = worksheet.Comments.GetThreadedComments(1, 1);
        foreach (ThreadedComment tc in b2Comments)
        {
            Console.WriteLine($"B2 comment: {tc.Notes} (Author: {tc.Author.Name})");
        }

        // Save the modified workbook
        workbook.Save("Output.xlsx");
    }
}