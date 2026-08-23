// Title: Count threaded comment occurrences per author in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate a C# function that accepts a Worksheet object and returns a Dictionary<string,int> mapping each threaded comment author to the number of comments they posted, using Aspose.Cells. | Extend the example to also tally top‑level (non‑threaded) comments, combine the totals per author, and display the results sorted by count descending. | Build a reusable utility class that scans all worksheets in a workbook, aggregates threaded comment authors across the entire file, and writes the author‑wise counts to a CSV file.
// Common Searches: aspocells c# count threaded comments by author in excel workbook | how to aggregate Excel comment authors using Aspose.Cells .NET | C# iterate through worksheet threaded comments and group by author Aspose.Cells | retrieve author statistics from Excel comment threads with Aspose.Cells for .NET
// Tags: Aspose.Cells threaded comment aggregation | C# author comment count dictionary | worksheet comment iteration Aspose.Cells | Excel comment thread analysis .NET | extract comment authors Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace ThreadedCommentAuthorCount
{
    // The sample loads an Excel workbook, iterates through each comment's threaded comments on the first worksheet, counts how many threaded comments each author has made using a case‑insensitive dictionary, prints the per‑author totals, and saves the workbook unchanged.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Get the first worksheet (or iterate through all worksheets as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Dictionary to hold author name and comment count
            Dictionary<string, int> authorCommentCount = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            // Iterate through all comments in the worksheet
            CommentCollection comments = worksheet.Comments;
            for (int i = 0; i < comments.Count; i++)
            {
                Comment comment = comments[i];

                // Get the threaded comments for this comment (if any)
                ThreadedCommentCollection threadedComments = comment.ThreadedComments;
                for (int j = 0; j < threadedComments.Count; j++)
                {
                    ThreadedComment tc = threadedComments[j];
                    string authorName = tc.Author?.Name ?? "Unknown";

                    // Increment count for this author
                    if (authorCommentCount.ContainsKey(authorName))
                        authorCommentCount[authorName]++;
                    else
                        authorCommentCount[authorName] = 1;
                }
            }

            // Output the results
            Console.WriteLine("Threaded comment count per author:");
            foreach (var kvp in authorCommentCount)
            {
                Console.WriteLine($"{kvp.Key}: {kvp.Value}");
            }

            // Optionally save the workbook (unchanged) to a new file
            workbook.Save("output.xlsx");
        }
    }
}
