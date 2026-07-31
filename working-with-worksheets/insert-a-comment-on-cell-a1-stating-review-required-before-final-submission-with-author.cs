// Title: Add a reviewer comment with author to cell A1 in Excel using Aspose.Cells for .NET (C#)
// Description: This C# example creates a new Workbook, accesses the first Worksheet, inserts a comment into cell A1, assigns "Reviewer Name" as the author and "Review required before final submission" as the note, and saves the file as CommentedWorkbook.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells add comment C# | Excel comment author .NET | C# insert cell note Aspose | programmatically add Excel comment | save workbook with comments Aspose.Cells
// Common Searches: how to add a comment with author to a cell using Aspose.Cells | Aspose.Cells C# set comment text and author | add review note to Excel cell programmatically | save Excel file with comments Aspose .NET
// Developer Intent: Insert a comment with a custom author into cell A1 and persist the workbook.
// Use Cases: Embed reviewer remarks directly into generated financial reports. | Automate audit annotations for compliance documentation. | Flag cells that need verification during data‑validation pipelines.
// AI Prompts: Write C# code with Aspose.Cells that adds a comment to cell B2, sets the author to "QA Lead", and changes the comment box background color. | Show how to add multiple comments to different cells, each with a unique author, using Aspose.Cells for .NET. | Demonstrate retrieving existing comments, updating their text, and re‑saving the workbook with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsCommentExample
{
    // This C# example creates a new Workbook, accesses the first Worksheet, inserts a comment into cell A1, assigns "Reviewer Name" as the author and "Review required before final submission" as the note, and saves the file as CommentedWorkbook.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a comment to cell A1 using the CommentCollection.Add(string) rule
            int commentIndex = worksheet.Comments.Add("A1");
            Comment comment = worksheet.Comments[commentIndex];

            // Set the comment author and note as required
            comment.Author = "Reviewer Name";
            comment.Note = "Review required before final submission";

            // Save the workbook (lifecycle save)
            workbook.Save("CommentedWorkbook.xlsx");
        }
    }
}
