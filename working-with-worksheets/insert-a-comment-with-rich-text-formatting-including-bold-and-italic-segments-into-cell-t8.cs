// Title: Insert bold and italic rich‑text comment into cell T8 using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, access the first worksheet, add a comment to cell T8, apply bold and italic formatting via the HtmlNote property, make the comment visible, and save the file as CommentRichTextDemo.xlsx with Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# comment HTML | rich text comment Aspose.Cells | cell T8 comment bold italic | Aspose.Cells HtmlNote | make comment visible Aspose.Cells | save workbook with comment Aspose.Cells | Aspose.Cells for .NET comment formatting | Excel comment rich text C#
// Common Searches: how to add a bold italic comment in Aspose.Cells | Aspose.Cells set HtmlNote for a cell comment | make Excel comment visible by default using Aspose.Cells | C# Aspose.Cells add formatted comment to specific cell | rich‑text comment example Aspose.Cells
// Developer Intent: Add a comment to cell T8 that contains bold and italic text, ensure the comment is visible, and save the workbook.
// Use Cases: Provide styled explanatory notes in financial models. | Highlight key cells for reviewers with bold/italic comments. | Embed in‑sheet documentation directly on important cells.
// AI Prompts: Generate C# code with Aspose.Cells to insert a comment containing bold and italic HTML into cell T8 and save the workbook. | Show how to extend the comment formatting to include underline, font color, or custom fonts using Aspose.Cells. | Explain how to programmatically toggle the visibility of comments after a workbook has been created with Aspose.Cells.

using Aspose.Cells;

namespace AsposeCellsCommentRichTextDemo
{
    // Demonstrates how to create a workbook, access the first worksheet, add a comment to cell T8, apply bold and italic formatting via the HtmlNote property, make the comment visible, and save the file as CommentRichTextDemo.xlsx with Aspose.Cells for .NET.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a comment to cell T8
            int commentIndex = worksheet.Comments.Add("T8");
            Comment comment = worksheet.Comments[commentIndex];

            // Set the comment text with bold and italic formatting using HTML
            comment.HtmlNote = "<b>Bold segment</b> and <i>Italic segment</i>";

            // Make the comment visible
            comment.IsVisible = true;

            // Save the workbook
            workbook.Save("CommentRichTextDemo.xlsx");
        }
    }
}
