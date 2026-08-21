// Title: Aspose.Cells .NET – Retrieve Threaded Comment Creation Time
// Description: Demonstrates how to add a threaded comment to a worksheet, obtain its CreatedTime property, and write the timestamp to the console using Aspose.Cells for .NET. The workbook is optionally saved to illustrate the full lifecycle.
// Keywords: Aspose.Cells | C# | .NET | ThreadedComment | CreatedTime | comment timestamp | Excel API | retrieve comment date
// Common Searches: Aspose.Cells get threaded comment created time | C# read CreatedTime of Excel comment | log comment timestamp with Aspose.Cells | how to access threaded comment date in .NET
// Developer Intent: Extract the creation timestamp of a threaded comment and display it.
// Use Cases: Audit when comments were added to a spreadsheet. | Show comment dates in a custom UI or report. | Validate comment age before processing.
// AI Prompts: Convert the CreatedTime of a threaded comment to ISO‑8601 format in Aspose.Cells. | Write C# code that compares the CreatedTime of two threaded comments and returns the newer one. | Create a method that iterates all threaded comments in a worksheet and writes their creation times to a log file.

using System;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentCreatedTimeDemo
{
    // Demonstrates how to add a threaded comment to a worksheet, obtain its CreatedTime property, and write the timestamp to the console using Aspose.Cells for .NET. The workbook is optionally saved to illustrate the full lifecycle.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a threaded comment author (or get an existing one)
            int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add("User1", "user1@example.com", "PROVIDER_1");
            ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

            // Add a threaded comment to cell A1
            int commentIndex = worksheet.Comments.Add("A1");
            Comment comment = worksheet.Comments[commentIndex];
            int threadedCommentIndex = comment.ThreadedComments.Add("Sample threaded comment", author);
            ThreadedComment threadedComment = comment.ThreadedComments[threadedCommentIndex];

            // Retrieve and log the creation time of the threaded comment
            DateTime createdTime = threadedComment.CreatedTime;
            Console.WriteLine("Threaded comment created at: " + createdTime);

            // Save the workbook (optional, just to complete the lifecycle)
            workbook.Save("ThreadedCommentCreatedTimeDemo.xlsx");
        }
    }
}
