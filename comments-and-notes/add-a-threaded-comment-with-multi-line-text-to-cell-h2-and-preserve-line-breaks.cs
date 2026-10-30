// Title: Add a threaded multi‑line comment with line breaks to cell H2 using Aspose.Cells for .NET
// Common Searches: how to insert a threaded comment with multiple lines into cell H2 using Aspose.Cells C# | preserving CRLF line breaks in Excel comments with Aspose.Cells .NET | setting comment author and making a comment visible in Aspose.Cells workbook | Aspose.Cells C# add comment to specific cell and keep line breaks | create Excel file with threaded comments programmatically in .NET
// Tags: add threaded comment to Excel cell Aspose.Cells | multi-line comment with line breaks C# | set comment author visibility Aspose.Cells | write comment to cell H2 .NET | create workbook with threaded comments Aspose | preserve CRLF in Excel comment Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, accesses the first worksheet, and adds a threaded comment to cell H2 containing multi‑line text with CRLF line breaks. It sets the comment author, makes the comment visible, and saves the workbook as 'ThreadedCommentExample.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Multi-line comment text with line breaks preserved
            string commentText = "First line of comment\r\nSecond line of comment\r\nThird line";

            // Add a comment to cell H2 (returns the comment index)
            int commentIndex = sheet.Comments.Add("H2");
            Comment comment = sheet.Comments[commentIndex];
            comment.Author = "UserName";
            comment.Note = commentText;
            comment.IsVisible = true; // Make the comment visible (optional)

            // Define output file path
            string outputPath = "ThreadedCommentExample.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
