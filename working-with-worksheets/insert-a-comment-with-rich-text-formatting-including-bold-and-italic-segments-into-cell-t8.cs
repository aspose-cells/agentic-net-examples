// Title: Aspose.Cells for .NET – Insert a Rich‑Text Comment (Bold & Italic) into cell T8 (C#)
// Description: Creates a new workbook, accesses the first worksheet, adds a comment to cell T8, applies HTML‑styled text with bold and italic segments via the HtmlNote property, makes the comment visible, and saves the file as an .xlsx document.
// Keywords: Aspose.Cells C# comment HtmlNote | rich text Excel comment .NET | bold italic comment Aspose.Cells | add comment to cell T8 | Excel comment formatting C# | Aspose.Cells tutorial
// Common Searches: Aspose.Cells add HTML formatted comment C# | How to set bold and italic text in an Excel comment using Aspose.Cells | Make Excel comment visible with Aspose.Cells .NET | Insert comment into specific cell T8 programmatically
// Developer Intent: Add a comment to cell T8 that contains both bold and italic text using Aspose.Cells for .NET.
// Use Cases: Annotate financial statements with emphasized notes. | Highlight assumptions in data‑model worksheets. | Automate generation of documentation comments during workbook creation.
// AI Prompts: Generate C# code that adds a comment with mixed bold and italic formatting to cell T8 using Aspose.Cells. | Show how to use the HtmlNote property to style an Excel comment and keep it visible. | Explain the difference between HtmlNote and Note properties for comment text in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsCommentRichTextDemo
{
    // Creates a new workbook, accesses the first worksheet, adds a comment to cell T8, applies HTML‑styled text with bold and italic segments via the HtmlNote property, makes the comment visible, and saves the file as an .xlsx document.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a comment to cell T8 (row 7, column 19)
            int commentIndex = worksheet.Comments.Add("T8");
            Comment comment = worksheet.Comments[commentIndex];

            // Set rich formatted text using HTML: bold and italic segments
            comment.HtmlNote = "<b>Bold segment</b> and <i>Italic segment</i> in the comment.";

            // Make the comment visible (optional)
            comment.IsVisible = true;

            // Save the workbook
            workbook.Save("CommentRichTextDemo.xlsx");
        }
    }
}
