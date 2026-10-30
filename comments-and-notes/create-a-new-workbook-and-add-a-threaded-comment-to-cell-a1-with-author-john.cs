// Title: Add a threaded comment authored by John to cell A1 in a new Excel workbook using Aspose.Cells for .NET
// Common Searches: how to add a threaded comment to cell A1 with Aspose.Cells for .NET | Aspose.Cells C# set comment author to John | create a new Excel workbook and insert a comment programmatically | save an Excel file with threaded comments using Aspose.Cells | Aspose.Cells add comment to specific cell example
// Tags: add threaded comment Aspose.Cells C# | set comment author Aspose.Cells | create workbook default worksheet Aspose.Cells | save workbook as xlsx Aspose.Cells | comment.Note property Aspose.Cells C#

using System;
using Aspose.Cells;

// The example creates a new Workbook, accesses the first worksheet, adds a threaded comment to cell A1, sets the author to "John" and the note text, then saves the file as ThreadedComment.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook with a default worksheet.
            Workbook workbook = new Workbook();

            // Get reference to the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Add a regular comment to cell A1.
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];

            // Set the author and text of the comment.
            comment.Author = "John";
            comment.Note = "This is a threaded comment.";

            // Save the workbook.
            string outputPath = "ThreadedComment.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
