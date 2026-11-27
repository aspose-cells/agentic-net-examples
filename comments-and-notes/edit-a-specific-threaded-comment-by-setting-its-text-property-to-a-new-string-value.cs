// Title: Update the Notes of a specific threaded comment in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# change threaded comment text in existing Excel file | how to edit a threaded comment's Notes property with Aspose.Cells .NET | programmatically modify Excel comment thread using Aspose.Cells API | C# set new text for a specific threaded comment in a worksheet | update first threaded comment in Excel workbook Aspose.Cells example
// Tags: Aspose.Cells edit threaded comment notes | C# update Excel comment thread text | modify threaded comment in workbook with Aspose.Cells | save workbook after comment edit Aspose.Cells | access comment thread collection C# Aspose.Cells

using System;
using Aspose.Cells;

namespace ThreadedCommentEditor
{
    // The sample loads an existing workbook (or creates one), accesses the first comment's first threaded comment, changes its Notes property to a new string, and saves the workbook to a new file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for the workbook
            const string workbookPath = "ThreadedCommentDemo.xlsx";
            const string outputPath = "ThreadedCommentDemo_Edited.xlsx";

            // Create a workbook instance. If the file already exists, load it; otherwise create a new one.
            Workbook workbook;
            if (System.IO.File.Exists(workbookPath))
            {
                Console.WriteLine($"Loading existing workbook: {workbookPath}");
                workbook = new Workbook(workbookPath);
            }
            else
            {
                Console.WriteLine("Workbook not found. Creating a new workbook with a threaded comment.");
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];

                // Add a comment to cell A1
                int commentIdx = ws.Comments.Add("A1");
                Comment comment = ws.Comments[commentIdx];
                comment.Note = "Base comment for A1";

                // Ensure there is at least one threaded comment author
                ThreadedCommentAuthorCollection authors = workbook.Worksheets.ThreadedCommentAuthors;
                int authorIdx = authors.Add("Demo Author", "demo_user", "demo_provider");
                ThreadedCommentAuthor author = authors[authorIdx];

                // Add a threaded comment to the comment we just created
                int threadedIdx = comment.ThreadedComments.Add("Original threaded comment text", author);
                ThreadedComment threadedComment = comment.ThreadedComments[threadedIdx];
                Console.WriteLine($"Created threaded comment with index {threadedIdx} and text: \"{threadedComment.Notes}\"");

                // Save the initial workbook so we have a file to edit later if needed
                workbook.Save(workbookPath);
                Console.WriteLine($"Initial workbook saved as {workbookPath}");
            }

            // At this point we have a workbook loaded with at least one threaded comment.
            // We'll edit the first threaded comment of the first comment in the first worksheet.
            try
            {
                Worksheet worksheet = workbook.Worksheets[0];

                if (worksheet.Comments.Count == 0)
                {
                    Console.WriteLine("No comments found in the worksheet. Exiting.");
                    return;
                }

                // Get the first comment (you can change the index to target a specific comment)
                Comment targetComment = worksheet.Comments[0];

                if (targetComment.ThreadedComments.Count == 0)
                {
                    Console.WriteLine("The selected comment does not contain any threaded comments. Exiting.");
                    return;
                }

                // Get the first threaded comment (change the index if you need a different one)
                ThreadedComment threaded = targetComment.ThreadedComments[0];

                // Display original text
                Console.WriteLine($"Original threaded comment text: \"{threaded.Notes}\"");

                // Set new text
                string newText = "Updated threaded comment text via Aspose.Cells";
                threaded.Notes = newText;

                // Verify the change
                Console.WriteLine($"Threaded comment text after update: \"{threaded.Notes}\"");

                // Save the workbook with the edited comment
                workbook.Save(outputPath);
                Console.WriteLine($"Edited workbook saved as {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while editing the threaded comment: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
