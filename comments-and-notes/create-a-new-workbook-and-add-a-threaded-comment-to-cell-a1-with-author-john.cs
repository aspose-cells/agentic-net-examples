// Title: Add a Threaded Comment by John to Cell A1 in a New Workbook – Aspose.Cells for .NET
// Description: C# example that creates a fresh Workbook, registers a threaded comment author named John, inserts a threaded comment with custom text into cell A1, and saves the file as ThreadedCommentDemo.xlsx using Aspose.Cells.
// Keywords: Aspose.Cells threaded comment C# | add comment to cell A1 | threaded comment author | create workbook Aspose.Cells | save Excel with comments .NET | Aspose.Cells API example
// Common Searches: how to add a threaded comment in Aspose.Cells C# | set author for threaded comments Aspose.Cells | save workbook with threaded comments .NET | Aspose.Cells add comment to specific cell
// Developer Intent: Insert a threaded comment authored by John into cell A1 of a newly created workbook.
// Use Cases: Annotate key cells in generated reports for reviewer feedback. | Maintain an audit trail by embedding author information in spreadsheet comments. | Automate collaborative note‑taking before distributing Excel files.
// AI Prompts: Show how to add multiple threaded comments with different authors in Aspose.Cells for .NET. | Provide code to list all threaded comments and their authors from an existing Excel workbook. | Explain how to modify the text of an existing threaded comment using Aspose.Cells.

using System;
using Aspose.Cells;

// C# example that creates a fresh Workbook, registers a threaded comment author named John, inserts a threaded comment with custom text into cell A1, and saves the file as ThreadedCommentDemo.xlsx using Aspose.Cells.
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
        workbook.Save("ThreadedCommentDemo.xlsx");
    }
}
