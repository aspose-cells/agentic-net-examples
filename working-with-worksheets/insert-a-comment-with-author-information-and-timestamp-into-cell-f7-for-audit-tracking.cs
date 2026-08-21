// Title: Add an audit comment with author and timestamp to cell F7 using Aspose.Cells for .NET
// Description: Demonstrates how to create a new workbook, insert a visible comment into cell F7, set the comment's author, embed the current date‑time (yyyy‑MM‑dd HH:mm:ss) and a custom audit note, then save the file as CommentWithAuditInfo.xlsx. Ideal for building an audit trail directly in Excel worksheets.
// Keywords: Aspose.Cells add comment C# | Excel audit comment timestamp | set comment author Aspose.Cells | visible comment cell F7 | save workbook with comment | audit trail Excel .NET | Aspose.Cells comment visibility
// Common Searches: Aspose.Cells add comment with date and author | C# insert audit comment into Excel cell | How to make a comment visible in Aspose.Cells | Add timestamp to Excel comment using .NET | Create audit trail in Excel with Aspose.Cells
// Developer Intent: Insert a visible comment that records the reviewer’s name and the current timestamp into cell F7 for audit tracking.
// Use Cases: Log review actions directly in the worksheet for compliance reporting. | Mark validated cells with reviewer identity and time of validation. | Provide an immutable audit trail in financial or regulatory Excel reports.
// AI Prompts: Generate C# code that uses Aspose.Cells to add a visible comment with author "John Doe" and the current timestamp to cell F7, then save the workbook. | Show how to update an existing comment in cell F7 to append a new timestamp each time the file is saved with Aspose.Cells for .NET. | Explain how to customize the comment note format, author name, and visibility settings for audit purposes in an Aspose.Cells workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsCommentExample
{
    // Demonstrates how to create a new workbook, insert a visible comment into cell F7, set the comment's author, embed the current date‑time (yyyy‑MM‑dd HH:mm:ss) and a custom audit note, then save the file as CommentWithAuditInfo.xlsx. Ideal for building an audit trail directly in Excel worksheets.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a comment to cell F7 using the CommentCollection.Add(string) method (feature rule)
            int commentIndex = worksheet.Comments.Add("F7");
            Comment comment = worksheet.Comments[commentIndex];

            // Set author information
            comment.Author = "John Doe";

            // Build the comment note with author and current timestamp
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            comment.Note = $"Author: {comment.Author}\nTimestamp: {timestamp}\nThis cell was reviewed.";

            // Optionally make the comment visible
            comment.IsVisible = true;

            // Save the workbook (lifecycle rule)
            workbook.Save("CommentWithAuditInfo.xlsx");
        }
    }
}
