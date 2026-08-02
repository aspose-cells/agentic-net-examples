// Title: Create an Excel workbook with threaded comments from multiple authors using Aspose.Cells (C#)
// Description: Shows how to instantiate a Workbook, add two ThreadedCommentAuthor objects, place a comment thread in B2 and another comment in C3, then save the file as ThreadedCommentsDemo.xlsx with Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# example | threaded comments Excel | add ThreadedCommentAuthor | worksheet comment thread | save workbook as xlsx | Excel collaboration Aspose | ThreadedComment overload | Aspose.Cells .NET tutorial | multiple authors comments
// Common Searches: Aspose.Cells add threaded comment C# | how to create comment thread in Excel with Aspose | multiple authors threaded comments Aspose.Cells | save workbook with comments Aspose .NET | threaded comment cell name overload Aspose
// Developer Intent: Insert threaded comment threads from several authors into specific cells and export the workbook as an XLSX file.
// Use Cases: Enable collaborative review by attaching a conversation thread to a cell. | Provide pre‑filled guidance comments in a template for data entry teams. | Record an audit trail of reviewer feedback directly within calculation cells.
// AI Prompts: Generate C# code that reads all threaded comments from a worksheet and prints author, text, and timestamp using Aspose.Cells. | Show how to modify the text of an existing threaded comment and reassign its author in a saved workbook with Aspose.Cells for .NET. | Provide an example that exports all threaded comments from an Excel file to a JSON array using Aspose.Cells.

using System;
using Aspose.Cells;

namespace ThreadedCommentsDemo
{
    // Shows how to instantiate a Workbook, add two ThreadedCommentAuthor objects, place a comment thread in B2 and another comment in C3, then save the file as ThreadedCommentsDemo.xlsx with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add threaded comment authors to the workbook
            int authorAliceIndex = workbook.Worksheets.ThreadedCommentAuthors.Add("Alice", "alice@example.com", "PROV1");
            ThreadedCommentAuthor authorAlice = workbook.Worksheets.ThreadedCommentAuthors[authorAliceIndex];

            int authorBobIndex = workbook.Worksheets.ThreadedCommentAuthors.Add("Bob", "bob@example.com", "PROV2");
            ThreadedCommentAuthor authorBob = workbook.Worksheets.ThreadedCommentAuthors[authorBobIndex];

            // Add threaded comments to cell B2 using the cell name overload
            worksheet.Comments.AddThreadedComment("B2", "Initial comment by Alice.", authorAlice);
            worksheet.Comments.AddThreadedComment("B2", "Reply from Bob.", authorBob);

            // Add a threaded comment to cell C3 using the row/column overload (row=2, column=2)
            worksheet.Comments.AddThreadedComment(2, 2, "Another comment on C3 by Alice.", authorAlice);

            // Save the workbook in XLSX format
            workbook.Save("ThreadedCommentsDemo.xlsx");
        }
    }
}
