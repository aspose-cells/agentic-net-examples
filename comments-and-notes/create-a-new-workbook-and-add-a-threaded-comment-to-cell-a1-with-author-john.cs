// Title: Create a new workbook and add a threaded comment authored by John to cell A1 using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells add a threaded comment to cell A1 | how to set a threaded comment author in an Aspose.Cells workbook | Aspose.Cells .NET create new Excel file with threaded notes | programmatically add a threaded comment with author John using Aspose.Cells | save workbook after inserting a threaded comment with Aspose.Cells C#
// Tags: add threaded comment to cell Aspose.Cells C# | threaded comment author configuration Aspose.Cells | create new workbook with threaded notes .NET | Aspose.Cells save workbook as xlsx with comments | Worksheet.Comments AddThreadedComment API

using System;
using Aspose.Cells;

// The example demonstrates how to instantiate a Workbook, define a threaded comment author named John, add a threaded comment to cell A1, display the comment details, and save the workbook as an .xlsx file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a threaded comment author named "John"
            int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add("John", "", "");
            ThreadedCommentAuthor johnAuthor = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

            // Add a threaded comment to cell A1 (row 0, column 0)
            worksheet.Comments.AddThreadedComment(0, 0, "This is a threaded comment by John.", johnAuthor);

            // Retrieve the comment to demonstrate that it was added
            Comment comment = worksheet.Comments[0];
            ThreadedComment threadedComment = comment.ThreadedComments[0];

            // Get the cell name from row and column indices
            string cellName = CellsHelper.CellIndexToName(comment.Row, comment.Column);

            Console.WriteLine($"Threaded comment added to {cellName}:");
            Console.WriteLine($"  Author: {threadedComment.Author.Name}");
            Console.WriteLine($"  Text  : {threadedComment.Notes}");

            // Save the workbook
            string outputPath = "ThreadedCommentDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while creating the threaded comment:");
            Console.WriteLine(ex.Message);
        }
    }
}
