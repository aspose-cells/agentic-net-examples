// Title: C# – Count All Threaded Comments in an Aspose.Cells Workbook
// Description: Creates a workbook, adds a threaded‑comment author, inserts several threaded comments, then walks through every worksheet and each comment’s ThreadedComments collection to compute and display the total number of threaded comments before saving the file.
// Keywords: Aspose.Cells | threaded comments | C# | count comments | Workbook API | Excel comment collection | ThreadedCommentAuthors | Comment.ThreadedComments
// Common Searches: Aspose.Cells count threaded comments .NET | How to get total threaded comments in Excel using Aspose.Cells C# | C# iterate worksheets to sum threaded comments Aspose | Retrieve threaded comment count from workbook Aspose.Cells
// Developer Intent: Find the total number of threaded comments that exist across all worksheets in a workbook.
// Use Cases: Ensure a workbook stays within a comment‑limit policy before distribution. | Produce an audit report showing comment activity per sheet for compliance reviews. | Identify sheets with the most discussion threads to focus quality‑control efforts.
// AI Prompts: Generate C# code that opens an Aspose.Cells workbook and returns the sum of all threaded comments. | Explain how to access the ThreadedComments collection for each comment and aggregate the counts in Aspose.Cells for .NET. | Show an example that groups threaded comments by author and prints a per‑author count using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentCount
{
    // Creates a workbook, adds a threaded‑comment author, inserts several threaded comments, then walks through every worksheet and each comment’s ThreadedComments collection to compute and display the total number of threaded comments before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a threaded comment author
            int authorIdx = workbook.Worksheets.ThreadedCommentAuthors.Add("John Doe", "john.doe@example.com", "JD");
            ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIdx];

            // Add some threaded comments to different cells
            worksheet.Comments.AddThreadedComment("A1", "First comment", author);
            worksheet.Comments.AddThreadedComment("B2", "Second comment", author);
            worksheet.Comments.AddThreadedComment("C3", "Third comment", author);

            // Calculate total number of threaded comments in the entire workbook
            int totalThreadedComments = 0;
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Iterate through each comment in the worksheet
                foreach (Comment comment in ws.Comments)
                {
                    // Each comment may contain a collection of threaded comments
                    totalThreadedComments += comment.ThreadedComments.Count;
                }
            }

            // Display the result
            Console.WriteLine($"Total threaded comments in the workbook: {totalThreadedComments}");

            // Save the workbook (lifecycle: save)
            workbook.Save("ThreadedCommentsCount.xlsx");
        }
    }
}
