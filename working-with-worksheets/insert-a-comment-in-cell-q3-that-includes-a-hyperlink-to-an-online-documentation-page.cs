// Title: Add a comment with a hyperlink to cell Q3 using Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, inserts a comment authored by "Automation" into cell Q3, attaches a hyperlink that points to the Aspose.Cells documentation with custom display text, and saves the file as CommentWithHyperlink.xlsx.
// Keywords: Aspose.Cells comment C# | Aspose.Cells hyperlink C# | add comment to cell Q3 | hyperlink with display text | save workbook Aspose.Cells
// Common Searches: Aspose.Cells add comment and hyperlink to same cell | C# insert comment with author in Aspose.Cells | how to set hyperlink display text in Aspose.Cells | save workbook with comment and link Aspose.Cells
// Developer Intent: Insert a comment in Q3, attach a documentation hyperlink with custom text, and write the workbook to disk.
// Use Cases: Provide inline help notes that link to online documentation. | Create reports where key cells contain both explanatory comments and direct resource links. | Automate spreadsheet generation with embedded guidance for end‑users.
// AI Prompts: Generate C# code that adds a comment and a hyperlink to cell Q3 with Aspose.Cells. | Explain how to add separate comments and hyperlinks to multiple cells in a single workbook. | Show how to configure an Aspose.Cells hyperlink to open in a new browser tab.

using System;
using Aspose.Cells;

// Creates a new workbook, inserts a comment authored by "Automation" into cell Q3, attaches a hyperlink that points to the Aspose.Cells documentation with custom display text, and saves the file as CommentWithHyperlink.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a comment to cell Q3 (row index 2, column index 16)
        int commentIdx = worksheet.Comments.Add("Q3");
        Comment comment = worksheet.Comments[commentIdx];
        comment.Author = "Automation";
        comment.Note = "See the online documentation for details.";

        // Add a hyperlink to the same cell Q3
        int hyperlinkIdx = worksheet.Hyperlinks.Add("Q3", 1, 1, "https://docs.aspose.com/cells/net/");
        // Set the text that will be displayed in the cell
        worksheet.Hyperlinks[hyperlinkIdx].TextToDisplay = "Aspose.Cells Documentation";

        // Save the workbook
        workbook.Save("CommentWithHyperlink.xlsx");
    }
}
