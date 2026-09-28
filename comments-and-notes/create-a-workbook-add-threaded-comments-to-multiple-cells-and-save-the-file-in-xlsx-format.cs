// Title: Generate an XLSX workbook with multiple threaded comments and authors using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a new workbook, defines two threaded comment authors, adds threaded comments to cells B2 and C3, and saves the file as an XLSX document. | Provide a C# snippet that enumerates the threaded comments of a given cell and prints each comment together with the author's name using Aspose.Cells.
// Common Searches: aspnet add threaded comments to specific cells with Aspose.Cells | how to assign multiple authors to threaded comments in an Excel file using C# | C# example to display comments attached to a single Excel cell using Aspose.Cells | save workbook with threaded comments as .xlsx using Aspose.Cells library | reply to a threaded comment in Aspose.Cells C#
// Tags: Aspose.Cells add threaded comment authors C# | Aspose.Cells create workbook with threaded comments | Aspose.Cells save workbook as XLSX with comments | Aspose.Cells retrieve threaded comments by cell | Aspose.Cells threaded comment reply example

using System;
using Aspose.Cells;

// The sample creates a new workbook, registers two threaded comment authors (Alice and Bob), adds threaded comments—including a reply—to cells B2 and C3, outputs the comments for B2, and saves the workbook as ThreadedCommentsDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add threaded comment authors
        int authorIndex1 = workbook.Worksheets.ThreadedCommentAuthors.Add("Alice", "alice@example.com", "PROV1");
        ThreadedCommentAuthor author1 = workbook.Worksheets.ThreadedCommentAuthors[authorIndex1];

        int authorIndex2 = workbook.Worksheets.ThreadedCommentAuthors.Add("Bob", "bob@example.com", "PROV2");
        ThreadedCommentAuthor author2 = workbook.Worksheets.ThreadedCommentAuthors[authorIndex2];

        // Add threaded comments to multiple cells using cell names
        worksheet.Comments.AddThreadedComment("B2", "First comment by Alice", author1);
        worksheet.Comments.AddThreadedComment("B2", "Reply by Bob", author2);
        worksheet.Comments.AddThreadedComment("C3", "Another comment by Alice", author1);

        // Optional: retrieve and display the threaded comments for verification
        ThreadedCommentCollection commentsB2 = worksheet.Comments.GetThreadedComments("B2");
        foreach (ThreadedComment tc in commentsB2)
        {
            Console.WriteLine($"Cell B2 comment: {tc.Notes} (by {tc.Author.Name})");
        }

        // Save the workbook in XLSX format
        workbook.Save("ThreadedCommentsDemo.xlsx");
    }
}
