// Title: Read and list the authors of threaded comments in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get author name of each threaded comment in Excel | list all comment authors from a worksheet using Aspose.Cells .NET | retrieve Excel comment author property with Aspose.Cells API | iterate through CommentCollection to output authors in C# | how to read threaded comment authors from .xlsx file using Aspose.Cells
// Tags: Aspose.Cells comment.Author property usage | C# enumerate worksheet CommentCollection | read threaded comment authors from Excel file | Aspose.Cells retrieve comment metadata .NET | list Excel comment authors programmatically

using Aspose.Cells;
using System;

// The example loads an Excel workbook, accesses a worksheet, iterates over its CommentCollection, and prints each comment's Author property, demonstrating how to extract threaded comment authors with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (or specify by name/index)
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the collection of comments in the worksheet
        CommentCollection comments = worksheet.Comments;

        // Iterate through each comment and output its author
        foreach (Comment comment in comments)
        {
            // For threaded comments, the Author property holds the comment author's name
            Console.WriteLine(comment.Author);
        }
    }
}
