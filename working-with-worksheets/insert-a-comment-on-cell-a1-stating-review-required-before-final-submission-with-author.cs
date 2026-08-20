// Title: C# – Add a comment with author to cell A1 using Aspose.Cells for .NET
// Description: Creates a new Workbook, accesses the first Worksheet, adds a comment to cell A1, sets the note to "Review required before final submission" and the author to "John Doe", then saves the file as CommentAdded.xlsx.
// Keywords: Aspose.Cells | Aspose.Cells for .NET | C# | add comment to Excel cell | cell comment author | Worksheet comment example | CommentCollection Add | Excel comment via Aspose | set comment text | save workbook with comment
// Common Searches: How to add a comment with author to a cell using Aspose.Cells C# | Aspose.Cells add note to A1 and set author | C# example for inserting Excel comments with Aspose | Save workbook after adding comments Aspose.Cells | Retrieve comment author from worksheet Aspose.Cells
// Developer Intent: Insert a comment with author information into cell A1 of an Excel worksheet using Aspose.Cells for .NET.
// Use Cases: Add reviewer notes to automatically generated reports for stakeholder feedback. | Highlight cells that need data validation or correction, attributing the comment to a specific author. | Create an audit trail by tagging key cells with author names before distribution.
// AI Prompts: Generate C# code to add a multi‑line, formatted comment to cell B2 with Aspose.Cells. | Show how to retrieve an existing comment from a worksheet and update its author using Aspose.Cells. | Provide a script that iterates through all comments in a worksheet and exports their text and authors to a CSV file.

using Aspose.Cells;

// Creates a new Workbook, accesses the first Worksheet, adds a comment to cell A1, sets the note to "Review required before final submission" and the author to "John Doe", then saves the file as CommentAdded.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (empty workbook)
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a comment to cell A1 using the CommentCollection.Add(string) overload
        int commentIndex = worksheet.Comments.Add("A1");
        Comment comment = worksheet.Comments[commentIndex];

        // Set the comment text
        comment.Note = "Review required before final submission";

        // Set the comment author
        comment.Author = "John Doe";

        // Save the workbook to a file
        workbook.Save("CommentAdded.xlsx");
    }
}
