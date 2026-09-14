// Title: How to create an Excel workbook and add a threaded comment to cell A1 with author John using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that initializes a new Workbook, registers a threaded comment author named John, inserts a threaded comment into cell A1, and saves the file as an .xlsx. | Demonstrate using the ThreadedCommentAuthors collection and the Comments.AddThreadedComment method to attach a comment to a specific cell in Aspose.Cells. | Show the steps to persist an Excel file after adding a threaded comment with a custom author in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# add threaded comment to cell A1 with custom author | How to set up ThreadedCommentAuthors in Aspose.Cells .NET | Saving an Excel workbook after inserting a threaded comment using Aspose.Cells | Example code for adding a threaded comment in Aspose.Cells for .NET | Create workbook and add comment author John Aspose.Cells C#
// Tags: Aspose.Cells threaded comment workflow | C# threaded comment author setup | Insert comment into specific cell Aspose.Cells | Save Excel file after comment insertion | Excel threaded comment author configuration .NET

using System;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentDemo
{
    // Creates a new workbook, registers a threaded comment author named John, adds a threaded comment to cell A1, and saves the workbook as ThreadedCommentDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the Aspose.Cells create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a threaded comment author named "John"
            // Parameters: name, userId, providerId
            int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add("John", "john@example.com", "PROVIDER");
            ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

            // Add a threaded comment to cell A1 (row 0, column 0) with the author
            worksheet.Comments.AddThreadedComment(0, 0, "This is a threaded comment by John.", author);

            // Save the workbook (uses the Aspose.Cells save rule)
            workbook.Save("ThreadedCommentDemo.xlsx");
        }
    }
}
