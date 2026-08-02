// Title: C# – Add an audit comment with author and timestamp to cell F7 using Aspose.Cells
// Description: Demonstrates how to create a new workbook, access the first worksheet, insert a visible comment into cell F7, set the author to "AuditSystem", embed the current date‑time (yyyy‑MM‑dd HH:mm:ss) in the note, and save the file as AuditComment.xlsx with Aspose.Cells for .NET.
// Keywords: Aspose.Cells add comment C# | timestamped Excel comment | audit comment Aspose.Cells | set comment author .NET | make comment visible Excel | save workbook with comment
// Common Searches: how to add a comment with author and date to a cell using Aspose.Cells | Aspose.Cells C# insert timestamped comment | make Excel comment visible programmatically | audit trail comment in Excel with Aspose.Cells | add comment to cell F7 Aspose.Cells
// Developer Intent: Insert a visible comment that records the author name and current timestamp into cell F7 for auditing purposes.
// Use Cases: Embed an audit trail directly in a worksheet by adding a timestamped comment to a key cell. | Provide data‑entry verification in financial or regulatory reports through visible author/date notes. | Automate compliance documentation by programmatically stamping cells with author and creation time.
// AI Prompts: Write C# code with Aspose.Cells that adds a visible comment to cell F7, sets the author, includes the current datetime, and saves the workbook. | Create a reusable method that accepts an author string, cell address, and workbook, then adds a timestamped comment using Aspose.Cells. | Explain how to format a datetime string for a comment note and ensure the comment remains visible in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AuditCommentExample
{
    // Demonstrates how to create a new workbook, access the first worksheet, insert a visible comment into cell F7, set the author to "AuditSystem", embed the current date‑time (yyyy‑MM‑dd HH:mm:ss) in the note, and save the file as AuditComment.xlsx with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a comment to cell F7 (row 7, column F)
            // Using the string overload for clarity
            int commentIndex = worksheet.Comments.Add("F7");
            Comment comment = worksheet.Comments[commentIndex];

            // Set author name
            comment.Author = "AuditSystem";

            // Build the comment note with timestamp
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            comment.Note = $"Record created by {comment.Author} on {timestamp}.";

            // Optionally make the comment visible
            comment.IsVisible = true;

            // Save the workbook
            workbook.Save("AuditComment.xlsx");
        }
    }
}
