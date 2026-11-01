// Title: Counting comments per author in an Excel worksheet with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells count how many comments each author has in an Excel sheet | retrieve comment author statistics from a workbook using Aspose.Cells | aggregate Excel comment authors with Aspose.Cells comment collection | load Excel file and tally comment authors in .NET
// Tags: Aspose.Cells comment author counting | C# iterate worksheet CommentCollection | Excel comment aggregation using Aspose.Cells | case-insensitive dictionary for author tally | load workbook and process comments Aspose

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, accesses the first worksheet's CommentCollection, and uses a case‑insensitive dictionary to count how many comments each author has made, then outputs the totals to the console.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        try
        {
            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Use regular comments collection (ThreadedComments may not be available in older versions)
            CommentCollection comments = sheet.Comments;

            // Dictionary to hold author name and comment count (case‑insensitive)
            Dictionary<string, int> authorCommentCount = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            // Iterate through all comments
            foreach (Comment comment in comments)
            {
                string author = comment.Author ?? "Unknown";

                if (authorCommentCount.ContainsKey(author))
                    authorCommentCount[author]++;
                else
                    authorCommentCount[author] = 1;
            }

            // Output the result
            Console.WriteLine("Comment count per author:");
            foreach (var kvp in authorCommentCount)
            {
                Console.WriteLine($"{kvp.Key}: {kvp.Value}");
            }

            // (Optional) Save the workbook if any changes were made
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while processing the worksheet: {ex.Message}");
        }
    }
}
