// Title: How to delete a threaded comment from cell C3 using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# remove threaded comment from a specific cell | delete threaded comment in an Excel workbook with Aspose.Cells .NET | remove comment thread from cell C3 programmatically Aspose.Cells | using ThreadedComments.RemoveAt to clear a threaded comment in C# | clear threaded discussion from a comment in an Aspose.Cells worksheet
// Tags: Aspose.Cells ThreadedComments.RemoveAt method | C# delete threaded comment in Excel workbook | remove threaded comment from worksheet cell | Aspose.Cells comment management API | Excel threaded comment removal C#

using System;
using Aspose.Cells;

// The example creates a workbook, adds a regular comment to cell C3, defines a threaded comment author, inserts a threaded comment, removes it with ThreadedComments.RemoveAt, and saves the file as "RemovedThreadedComment.xlsx".
class Program
{
    static void Main()
    {
        try
        {
            // Path to output file
            const string outputPath = "RemovedThreadedComment.xlsx";

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // 1. Add a regular comment to cell C3 (required to host threaded comments)
            // ------------------------------------------------------------
            int commentIndex = worksheet.Comments.Add("C3");
            Comment comment = worksheet.Comments[commentIndex];
            comment.Note = "Parent comment for threaded discussion";

            // ------------------------------------------------------------
            // 2. Add a threaded comment author to the workbook
            // ------------------------------------------------------------
            ThreadedCommentAuthorCollection authors = workbook.Worksheets.ThreadedCommentAuthors;
            int authorId = authors.Add("Demo Author", "demo_user@example.com", "DemoProvider");
            ThreadedCommentAuthor author = authors[authorId];

            // ------------------------------------------------------------
            // 3. Add a threaded comment to the comment we just created
            // ------------------------------------------------------------
            int threadedIndex = comment.ThreadedComments.Add("This is a threaded comment that will be removed", author);
            ThreadedComment threadedComment = comment.ThreadedComments[threadedIndex];

            Console.WriteLine($"Added threaded comment at index {threadedIndex} with text: \"{threadedComment.Notes}\"");

            // ------------------------------------------------------------
            // 4. Remove the threaded comment using the RemoveAt method on the collection
            // ------------------------------------------------------------
            comment.ThreadedComments.RemoveAt(threadedIndex);
            Console.WriteLine("Threaded comment removed successfully.");

            // ------------------------------------------------------------
            // 5. Save the workbook to verify the comment has been removed
            // ------------------------------------------------------------
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
