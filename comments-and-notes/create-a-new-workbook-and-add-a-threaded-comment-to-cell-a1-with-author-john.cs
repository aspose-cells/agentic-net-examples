// Title: Add a Threaded Comment by John to Cell A1 in a New Workbook – Aspose.Cells for .NET (C#)
// Description: This C# snippet creates a fresh Workbook with Aspose.Cells, registers a threaded‑comment author named John, inserts a threaded comment into cell A1, and saves the file as ThreadedCommentJohn.xlsx.
// Keywords: Aspose.Cells | C# | .NET | threaded comment | add threaded comment | comment author | Excel workbook | cell A1 comment | ThreadedCommentJohn.xlsx
// Common Searches: how to add a threaded comment in Aspose.Cells C# | Aspose.Cells set comment author | create workbook with threaded comment .NET | add comment to cell A1 using Aspose.Cells | threaded comment example C#
// Developer Intent: Insert a threaded comment authored by John into cell A1 of a newly created Excel workbook using Aspose.Cells for .NET.
// Use Cases: Add reviewer notes to specific cells in a generated report, preserving each reviewer’s identity. | Automate the insertion of audit‑trail comments into a template spreadsheet before distribution. | Enable collaborative editing by attaching author‑specific threaded comments to key data points.
// AI Prompts: Show how to add multiple threaded comments with different authors to a workbook using Aspose.Cells for .NET. | Provide code to list all threaded comments and their authors from an existing Excel file with Aspose.Cells. | Explain how to modify the text of an existing threaded comment and change its author in a saved workbook.

using System;
using Aspose.Cells;

// This C# snippet creates a fresh Workbook with Aspose.Cells, registers a threaded‑comment author named John, inserts a threaded comment into cell A1, and saves the file as ThreadedCommentJohn.xlsx.
class ThreadedCommentExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a threaded comment author named John
        int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add("John", "john@example.com", "DEFAULT");
        ThreadedCommentAuthor johnAuthor = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

        // Add a threaded comment to cell A1 using the author
        worksheet.Comments.AddThreadedComment("A1", "This is a threaded comment added by John.", johnAuthor);

        // Save the workbook
        workbook.Save("ThreadedCommentJohn.xlsx");
    }
}
