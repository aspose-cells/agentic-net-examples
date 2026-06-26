using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

class ThreadedCommentAuthorCount
{
    static void Main()
    {
        const string inputPath = "ThreadedComments.xlsx";
        const string outputPath = "ThreadedComments_Output.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the workbook containing threaded comments
            Workbook workbook = new Workbook(inputPath);

            // Store the number of threaded comments per author (case‑insensitive)
            Dictionary<string, int> authorCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get all comments on the current sheet
                CommentCollection comments = sheet.Comments;

                // Process each comment (attached to a cell)
                foreach (Comment comment in comments)
                {
                    // Retrieve threaded comments for the cell
                    ThreadedCommentCollection threadedComments = comment.ThreadedComments;

                    // Count each threaded comment by its author
                    foreach (ThreadedComment tc in threadedComments)
                    {
                        string authorName = tc.Author?.Name ?? "Unknown";

                        if (authorCounts.ContainsKey(authorName))
                            authorCounts[authorName]++;
                        else
                            authorCounts[authorName] = 1;
                    }
                }
            }

            // Output the results
            Console.WriteLine("Threaded comment count per author:");
            foreach (var kvp in authorCounts)
            {
                Console.WriteLine($"{kvp.Key}: {kvp.Value}");
            }

            // Save the workbook (no modifications made, optional)
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}