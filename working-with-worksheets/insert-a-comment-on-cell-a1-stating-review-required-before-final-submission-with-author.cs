// Title: Add a reviewer comment with author to cell A1 and save the workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a new Workbook, adds a comment authored by a specified name with the text "Review required before final submission" to cell A1, and saves the file as an .xlsx using Aspose.Cells. | Write a C# snippet that sets the Author and Note properties of a comment on cell A1, then persists the workbook to disk with Aspose.Cells. | Provide a C# example that demonstrates how to insert a single‑cell comment with custom author information and export the workbook to "CommentedWorkbook.xlsx" using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# how to add a comment with author to cell A1 | set comment author and note in Aspose.Cells workbook programmatically | save workbook after inserting cell comment using Aspose.Cells for .NET | C# Aspose.Cells add reviewer note to specific cell | example code for adding comments to Excel cells with Aspose.Cells .NET
// Tags: add comment with author Aspose.Cells .NET | cell A1 comment insertion Aspose.Cells | set comment note property C# Aspose.Cells | save workbook after comment Aspose.Cells | Aspose.Cells comment API usage

using Aspose.Cells;
using System;

// Creates a new workbook, adds a comment authored by "AuthorName" with the note "Review required before final submission" to cell A1 of the first worksheet, and saves the file as CommentedWorkbook.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet (default sheet)
            Worksheet sheet = workbook.Worksheets[0];

            // Add a comment to cell A1; Add returns the comment index
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];
            comment.Author = "AuthorName";
            comment.Note = "Review required before final submission";

            // Save the workbook to a file
            workbook.Save("CommentedWorkbook.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
