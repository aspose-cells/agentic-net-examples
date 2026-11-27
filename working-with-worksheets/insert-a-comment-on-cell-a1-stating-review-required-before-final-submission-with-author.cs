// Title: Insert a visible author comment into cell A1 of an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: how to add a visible comment with author to cell A1 using Aspose.Cells C# | Aspose.Cells set comment note and author for a specific Excel cell | make Excel comment visible when opening workbook with Aspose.Cells .NET | programmatically add author information to an Excel cell comment using Aspose.Cells | save workbook with visible comments using Aspose.Cells for .NET
// Tags: Aspose.Cells add comment C# | set comment author Excel .NET | make comment visible Aspose.Cells | Excel cell A1 comment Aspose.Cells | save workbook with comments .NET

using System;
using Aspose.Cells;

// The example creates a new workbook, accesses the first worksheet, adds a comment to cell A1 with the note "Review required before final submission" and author "John Doe", makes the comment visible on opening, and saves the file as CommentDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a comment to cell A1
            int commentIndex = worksheet.Comments.Add("A1");
            Comment comment = worksheet.Comments[commentIndex];

            // Set the comment text and author
            comment.Note = "Review required before final submission";
            comment.Author = "John Doe";

            // Make the comment visible when the file is opened
            comment.IsVisible = true;

            // Save the workbook
            string outputPath = "CommentDemo.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Comment successfully added to A1 and saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
