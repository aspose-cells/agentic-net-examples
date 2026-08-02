// Title: C# – Add a comment with a documentation hyperlink to cell Q3 using Aspose.Cells
// Description: Creates a new workbook, adds a visible comment to cell Q3 that contains a link to the Aspose.Cells .NET documentation, optionally inserts a clickable hyperlink on the same cell, and saves the file as CommentWithHyperlink.xlsx.
// Keywords: Aspose.Cells comment hyperlink C# | add comment to cell Q3 | Aspose.Cells insert hyperlink | C# Excel comment with URL | Aspose.Cells documentation link
// Common Searches: how to add a comment with a URL in Aspose.Cells C# | insert hyperlink and comment in the same Excel cell using Aspose.Cells | Aspose.Cells add comment to Q3 | C# Aspose.Cells add clickable hyperlink to a cell
// Developer Intent: Insert a visible comment in cell Q3 that includes a link to the Aspose.Cells documentation and optionally add a clickable hyperlink on the same cell.
// Use Cases: Provide instant access to online API docs from a spreadsheet cell. | Create tutorial workbooks where each step’s description is stored in a comment with a reference link. | Combine explanatory text (comment) and direct navigation (hyperlink) for end‑users.
// AI Prompts: Generate C# code with Aspose.Cells that adds a comment containing a documentation hyperlink to cell Q3 and saves the workbook. | Show how to style the comment and customize the hyperlink text for a cell using Aspose.Cells for .NET. | Explain how to read, update, or remove an existing comment’s hyperlink in an Aspose.Cells workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsCommentWithHyperlink
{
    // Creates a new workbook, adds a visible comment to cell Q3 that contains a link to the Aspose.Cells .NET documentation, optionally inserts a clickable hyperlink on the same cell, and saves the file as CommentWithHyperlink.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a comment to cell Q3
            CommentCollection comments = worksheet.Comments;
            int commentIndex = comments.Add("Q3");               // Add comment by cell name
            Comment comment = comments[commentIndex];
            comment.Author = "Aspose";
            comment.IsVisible = true;
            // Include the hyperlink URL in the comment text
            comment.Note = "For detailed documentation, visit: https://docs.aspose.com/cells/net/";

            // Optionally, also add a clickable hyperlink to the same cell
            int hyperlinkIndex = worksheet.Hyperlinks.Add("Q3", 1, 1, "https://docs.aspose.com/cells/net/");
            Hyperlink hyperlink = worksheet.Hyperlinks[hyperlinkIndex];
            hyperlink.TextToDisplay = "Aspose.Cells Documentation";

            // Save the workbook
            workbook.Save("CommentWithHyperlink.xlsx");
        }
    }
}
