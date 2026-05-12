using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace ThreadedCommentCounter
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // Dictionary to hold comment count per author name
            Dictionary<string, int> authorCommentCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Iterate through all comments in the worksheet
                foreach (Comment comment in worksheet.Comments)
                {
                    // Get the collection of threaded comments for this comment
                    ThreadedCommentCollection threadedComments = comment.ThreadedComments;

                    // Iterate through each threaded comment
                    foreach (ThreadedComment threadedComment in threadedComments)
                    {
                        // Get author name (fallback to "Unknown" if null)
                        string authorName = threadedComment.Author?.Name ?? "Unknown";

                        // Increment count for this author
                        if (authorCommentCounts.ContainsKey(authorName))
                            authorCommentCounts[authorName]++;
                        else
                            authorCommentCounts[authorName] = 1;
                    }
                }
            }

            // Display the results
            Console.WriteLine("Threaded comment count per author:");
            foreach (KeyValuePair<string, int> kvp in authorCommentCounts)
            {
                Console.WriteLine($"{kvp.Key}: {kvp.Value}");
            }

            // Save the workbook if any modifications were made (optional)
            workbook.Save("OutputWorkbook.xlsx");
        }
    }
}