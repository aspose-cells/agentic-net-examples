using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("Input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a threaded comment author to the workbook
        int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add(
            "Demo Author",          // Author name
            "demo@example.com",     // User ID / email
            "DemoProvider");        // Provider ID
        ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

        // Insert a threaded comment into cell B2 (row index 1, column index 1)
        worksheet.Comments.AddThreadedComment(1, 1, "This is a threaded comment.", author);

        // Retrieve and display the threaded comments for verification
        ThreadedCommentCollection comments = worksheet.Comments.GetThreadedComments(1, 1);
        foreach (ThreadedComment tc in comments)
        {
            Console.WriteLine($"Comment: {tc.Notes} (Author: {tc.Author.Name})");
        }

        // Save the workbook with the new threaded comment
        workbook.Save("Output.xlsx");
    }
}