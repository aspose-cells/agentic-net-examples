// Title: Aspose.Cells for .NET – Retrieve Threaded Comment Creation Time (CreatedTime)
// Description: Demonstrates how to create a workbook, add a threaded comment author, insert a threaded comment into cell A1, read the comment's CreatedTime property, output the timestamp to the console, and optionally save the file.
// Keywords: Aspose.Cells .NET threaded comment CreatedTime | get comment creation timestamp Aspose.Cells | ThreadedComment.CreatedTime example | C# read Excel comment date | audit comment timestamps Aspose | global Excel comment metadata
// Common Searches: Aspose.Cells how to read CreatedTime of a threaded comment | C# example for getting Excel comment creation date | Threaded comment timestamp Aspose.Cells .NET | Retrieve and display comment creation time in Excel workbook
// Developer Intent: Extract the creation timestamp of a threaded comment and display or log it.
// Use Cases: Maintain an audit trail of when comments were added to generated reports | Show comment timestamps in a custom dashboard or UI | Filter or flag comments older than a specific date for review
// AI Prompts: Write C# code using Aspose.Cells to fetch a threaded comment's CreatedTime and format it as ISO‑8601. | Show how to loop through all threaded comments in a worksheet and print each CreatedTime value. | Explain how to compare ThreadedComment.CreatedTime with the current date to identify recent comments.

using System;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentCreatedTimeDemo
{
    // Demonstrates how to create a workbook, add a threaded comment author, insert a threaded comment into cell A1, read the comment's CreatedTime property, output the timestamp to the console, and optionally save the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a threaded comment author (or get the default one)
            int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add("User1", "user1@example.com", "Provider1");
            ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

            // Add a threaded comment to cell A1
            int commentIndex = worksheet.Comments.Add("A1");
            Comment comment = worksheet.Comments[commentIndex];
            int threadedCommentIndex = comment.ThreadedComments.Add("Sample threaded comment", author);
            ThreadedComment threadedComment = comment.ThreadedComments[threadedCommentIndex];

            // Retrieve and log the creation time of the threaded comment
            DateTime createdTime = threadedComment.CreatedTime;
            Console.WriteLine("Threaded comment created at: " + createdTime);

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("ThreadedCommentCreatedTimeDemo.xlsx");
        }
    }
}
