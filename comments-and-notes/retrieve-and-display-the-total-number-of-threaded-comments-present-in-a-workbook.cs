// Title: Count all threaded comments in an Excel workbook using Aspose.Cells for C#
// AI Prompts: Write C# code that opens an Excel file with Aspose.Cells, loops through every worksheet and comment, and returns the total number of threaded comments. | Create a method in C# that accepts a Workbook object and calculates the sum of ThreadedComments across all cells, then prints the result. | Show how to modify the example to produce a dictionary of worksheet names mapped to their individual threaded comment counts using Aspose.Cells.
// Common Searches: C# Aspose.Cells how to get total threaded comment count in a workbook | sum of threaded comments across all sheets using Aspose.Cells API | retrieve number of comment threads per worksheet Aspose.Cells C# | example code to count Excel threaded comments with Aspose.Cells | Aspose.Cells count threaded comments in each cell C#
// Tags: count threaded comments Aspose.Cells C# | iterate workbook comment collection Aspose.Cells | threaded comment aggregation Excel C# | retrieve total comment threads Aspose.Cells | worksheet level threaded comment count Aspose.Cells

using System;
using Aspose.Cells;

namespace ThreadedCommentsCountDemo
{
    // Demonstrates creating a workbook, adding threaded comments, iterating through each worksheet and its comments to compute the total number of threaded comments, outputting the count, and optionally saving the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the collection of threaded comment authors
            ThreadedCommentAuthorCollection authors = workbook.Worksheets.ThreadedCommentAuthors;
            int authorIdx = authors.Add("John Doe", "john.doe@example.com", "JD");
            ThreadedCommentAuthor author = authors[authorIdx];

            // Add some threaded comments to different cells
            worksheet.Comments.AddThreadedComment("A1", "First comment in A1", author);
            worksheet.Comments.AddThreadedComment("A1", "Second comment in A1", author);
            worksheet.Comments.AddThreadedComment("B2", "Only comment in B2", author);
            worksheet.Comments.AddThreadedComment("C3", "First comment in C3", author);
            worksheet.Comments.AddThreadedComment("C3", "Second comment in C3", author);
            worksheet.Comments.AddThreadedComment("C3", "Third comment in C3", author);

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

            // Save the workbook (optional, just to demonstrate lifecycle usage)
            workbook.Save("ThreadedCommentsCountDemo.xlsx");
        }
    }
}
