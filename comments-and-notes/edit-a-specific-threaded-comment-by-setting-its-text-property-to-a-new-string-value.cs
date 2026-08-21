// Title: C# – Edit Threaded Comment Text in Excel with Aspose.Cells
// Description: Shows how to create a workbook, add a threaded comment to cell A1, retrieve the comment, modify its text using the ThreadedComment.Notes property, and save the result as EditedThreadedComment.xlsx with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | threaded comment | edit comment text | ThreadedComment.Notes | Excel workbook | programmatic comment update | modify Excel comment | Aspose.Cells API
// Common Searches: Aspose.Cells change threaded comment text C# | update ThreadedComment Notes property | edit Excel threaded comment programmatically | Aspose.Cells set comment notes .NET | how to modify threaded comment in workbook
// Developer Intent: Change the text of an existing threaded comment in an Excel worksheet using Aspose.Cells for .NET.
// Use Cases: Replace placeholder notes with final review comments after document approval. | Synchronize comment content with external data sources before publishing the workbook. | Batch‑correct spelling errors in threaded comments across multiple Excel files.
// AI Prompts: Generate C# code that finds a threaded comment by cell address and updates its Notes property with Aspose.Cells. | Provide an example that loops through all threaded comments in a worksheet and adds a timestamp prefix to each comment's text.

using System;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentEdit
{
    // Shows how to create a workbook, add a threaded comment to cell A1, retrieve the comment, modify its text using the ThreadedComment.Notes property, and save the result as EditedThreadedComment.xlsx with Aspose.Cells for .NET.
    public class EditThreadedComment
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Create or obtain a threaded comment author
                ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors["User1"];
                author.Name = "User1";

                // Add a threaded comment to cell A1 with initial text
                worksheet.Comments.AddThreadedComment(0, 0, "Initial comment text", author);

                // Retrieve the threaded comment that was just added
                Comment comment = worksheet.Comments[0];
                ThreadedComment threadedComment = comment.ThreadedComments[0];

                // Update the comment text using the Notes property
                threadedComment.Notes = "Updated comment text";

                // Optional: display the updated text to verify
                Console.WriteLine("Threaded comment updated to: " + threadedComment.Notes);

                // Save the workbook
                workbook.Save("EditedThreadedComment.xlsx");
                Console.WriteLine("Workbook saved as EditedThreadedComment.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            EditThreadedComment.Run();
        }
    }
}
