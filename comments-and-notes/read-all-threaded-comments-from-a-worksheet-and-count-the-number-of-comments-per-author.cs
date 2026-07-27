// Title: Count Excel Threaded Comments by Author with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, reads all threaded comments on a worksheet, extracts each comment's author, aggregates the number of comments per author, prints the totals, and saves the file. Demonstrates using CommentCollection, ThreadedCommentCollection, and a dictionary for tallying.
// Keywords: Aspose.Cells | C# | .NET | Excel threaded comments | comment author count | CommentCollection | ThreadedCommentCollection | read Excel comments | author statistics | Excel automation
// Common Searches: Aspose.Cells count threaded comments by author C# | how to get comment author statistics from Excel using Aspose | C# iterate Excel comments and tally per user | read threaded comments Aspose.Cells .NET | Excel comment author summary code
// Developer Intent: Read every threaded comment in a worksheet and calculate how many comments each author has contributed.
// Use Cases: Generate a reviewer contribution report for collaborative Excel workbooks. | Enforce comment‑limit policies before publishing a spreadsheet. | Create a dashboard that visualizes comment activity per team member.
// AI Prompts: Provide a reusable method that accepts a Worksheet and returns a Dictionary<string,int> of author comment counts using Aspose.Cells. | Modify the sample to ignore a specific author while still counting all other comments. | Add comprehensive error handling for missing files, empty comment collections, and log the author counts to a text file.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentCounter
{
    // Loads an Excel workbook, reads all threaded comments on a worksheet, extracts each comment's author, aggregates the number of comments per author, prints the totals, and saves the file. Demonstrates using CommentCollection, ThreadedCommentCollection, and a dictionary for tallying.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("ThreadedCommentsDemo.xlsx");

            // Access the first worksheet (or iterate through all worksheets if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the collection of comments in the worksheet
            CommentCollection comments = worksheet.Comments;

            // Dictionary to hold comment count per author name
            Dictionary<string, int> authorCommentCounts = new Dictionary<string, int>();

            // Iterate through each comment in the worksheet
            foreach (Comment comment in comments)
            {
                // Get the threaded comments associated with this comment
                ThreadedCommentCollection threadedComments = comment.ThreadedComments;

                // Iterate through each threaded comment
                foreach (ThreadedComment tc in threadedComments)
                {
                    // Retrieve the author name (fallback to "Unknown" if null)
                    string authorName = tc.Author?.Name ?? "Unknown";

                    // Increment the count for this author
                    if (authorCommentCounts.ContainsKey(authorName))
                        authorCommentCounts[authorName]++;
                    else
                        authorCommentCounts[authorName] = 1;
                }
            }

            // Output the results
            Console.WriteLine("Threaded comment count per author:");
            foreach (KeyValuePair<string, int> kvp in authorCommentCounts)
            {
                Console.WriteLine($"Author: {kvp.Key}, Comments: {kvp.Value}");
            }

            // Save the workbook (no modifications made, but required by lifecycle rule)
            workbook.Save("ThreadedCommentsCounted.xlsx");
        }
    }
}
