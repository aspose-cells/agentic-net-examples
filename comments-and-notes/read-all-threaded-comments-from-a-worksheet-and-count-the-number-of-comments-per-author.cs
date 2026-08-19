// Title: Aspose.Cells .NET Example – Count Threaded Comments per Author in Excel
// Description: Loads an Excel workbook, reads all threaded comments from the first worksheet, tallies each author's contributions with a case‑insensitive dictionary, prints the totals and optionally saves the file.
// Keywords: Aspose.Cells read threaded comments | C# count comment authors | Excel threaded comment example | Aspose.Cells comment author dictionary | C# Excel comment analysis
// Common Searches: how to count threaded comments Aspose.Cells | C# get comment author counts from Excel | Aspose.Cells enumerate worksheet comments | Excel comment author statistics C# | sample code for threaded comments Aspose
// Developer Intent: Extract every threaded comment from a worksheet and compute how many each author has posted.
// Use Cases: Create a report showing comment activity per collaborator for audit purposes. | Identify the most active reviewer before finalizing a shared spreadsheet. | Verify that all required stakeholders have left at least one comment.
// AI Prompts: Generate a method that returns a Dictionary<string,int> of author comment counts for a given Worksheet using Aspose.Cells. | Adapt the sample to ignore a specific author and display counts for the remaining participants. | Write unit tests that mock worksheets with threaded comments to validate the counting logic.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace ThreadedCommentCounter
{
    // Loads an Excel workbook, reads all threaded comments from the first worksheet, tallies each author's contributions with a case‑insensitive dictionary, prints the totals and optionally saves the file.
    class Program
    {
        static void Main()
        {
            const string inputPath = "InputWithThreadedComments.xlsx";
            const string outputPath = "OutputWithThreadedCommentsProcessed.xlsx";

            // Verify that the input workbook exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Dictionary to store comment counts per author (case‑insensitive)
                Dictionary<string, int> authorCommentCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                // Retrieve the comments collection from the worksheet
                CommentCollection comments = worksheet.Comments;

                // Iterate through each comment
                for (int i = 0; i < comments.Count; i++)
                {
                    Comment comment = comments[i];

                    // Get threaded comments associated with the current comment
                    ThreadedCommentCollection threadedComments = comment.ThreadedComments;

                    // Count each threaded comment by its author
                    foreach (ThreadedComment tc in threadedComments)
                    {
                        string authorName = tc.Author?.Name ?? "Unknown";

                        if (authorCommentCounts.ContainsKey(authorName))
                            authorCommentCounts[authorName]++;
                        else
                            authorCommentCounts[authorName] = 1;
                    }
                }

                // Output the results
                Console.WriteLine("Threaded comment count per author:");
                foreach (var kvp in authorCommentCounts)
                {
                    Console.WriteLine($"{kvp.Key}: {kvp.Value}");
                }

                // Save the workbook (optional)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
