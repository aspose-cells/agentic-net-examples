// Title: Add a comment to cell D4 in Excel with Aspose.Cells for .NET (C#)
// Description: C# code that creates a new workbook, accesses the first worksheet, inserts a comment into cell D4, sets the comment text to "Review required before final submission.", and saves the file as Output.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | add comment to Excel cell | cell D4 comment | Comments.Add | Comment.Note | save workbook Aspose.Cells | Output.xlsx | Excel annotation
// Common Searches: Aspose.Cells add comment to specific cell C# | How to set comment text for D4 using Aspose.Cells | Save Excel file after adding comments Aspose.Cells .NET | Insert review note in Excel cell with Aspose.Cells | C# Aspose.Cells comment API example
// Developer Intent: Insert a predefined comment into cell D4 of a newly created Excel workbook.
// Use Cases: Mark cells that need reviewer attention before publishing the spreadsheet. | Create an audit trail by programmatically adding notes to critical cells. | Automate data‑validation feedback by placing comments on cells that fail checks.
// AI Prompts: Generate C# code that adds a comment to cell D4 with Aspose.Cells and saves the workbook. | Show how to add multiple comments with different texts to various cells using Aspose.Cells for .NET. | Explain how to customize a comment’s author, font style, and background color after adding it to a cell with Aspose.Cells.

using System;
using Aspose.Cells;

// C# code that creates a new workbook, accesses the first worksheet, inserts a comment into cell D4, sets the comment text to "Review required before final submission.", and saves the file as Output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a comment to cell D4 using the cell name overload
        int commentIndex = worksheet.Comments.Add("D4");
        Comment comment = worksheet.Comments[commentIndex];

        // Set the comment text
        comment.Note = "Review required before final submission.";

        // Save the workbook to a file
        workbook.Save("Output.xlsx");
    }
}
