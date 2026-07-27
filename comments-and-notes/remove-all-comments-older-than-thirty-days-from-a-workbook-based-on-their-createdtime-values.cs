// Title: C# Sample: Delete Excel Comments Older Than 30 Days Using Aspose.Cells
// Description: This C# example loads an XLSX workbook with Aspose.Cells, computes a 30‑day cutoff, iterates each worksheet's CommentCollection, extracts the creation date from each comment's Note (or CreatedTime), removes comments older than the cutoff, and saves the cleaned workbook.
// Keywords: Aspose.Cells | C# | Excel comments | remove old comments | comment CreatedTime | filter comments by date | worksheet comment collection | Excel automation | sample code | GitHub example
// Common Searches: Aspose.Cells remove comments older than 30 days | C# delete Excel comments by date | how to filter worksheet comments in Aspose.Cells | parse comment note timestamp Aspose.Cells | remove stale Excel notes programmatically
// Developer Intent: Programmatically delete worksheet comments whose creation date exceeds 30 days.
// Use Cases: Clean up monthly financial reports by purging outdated reviewer notes before archiving. | Automate maintenance of shared spreadsheets in a corporate environment, ensuring only recent comments remain. | Prepare a workbook for public distribution by stripping comments that are older than a month.
// AI Prompts: Generate C# code with Aspose.Cells that removes Excel comments older than a specified number of days, handling missing or malformed dates gracefully. | Create a reusable method that parses a date from the beginning of a comment's Note property and returns true if the comment is older than a given cutoff. | Show how to iterate backwards through Worksheet.Comments to safely delete items in Aspose.Cells, including error handling and logging.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// This C# example loads an XLSX workbook with Aspose.Cells, computes a 30‑day cutoff, iterates each worksheet's CommentCollection, extracts the creation date from each comment's Note (or CreatedTime), removes comments older than the cutoff, and saves the cleaned workbook.
class RemoveOldComments
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Define the cutoff date (30 days ago)
            DateTime cutoffDate = DateTime.Now.AddDays(-30);

            // Process each worksheet
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                CommentCollection comments = worksheet.Comments;

                // Iterate backwards to safely remove items
                for (int i = comments.Count - 1; i >= 0; i--)
                {
                    Comment comment = comments[i];
                    bool remove = false;

                    if (!string.IsNullOrEmpty(comment.Note))
                    {
                        string[] parts = comment.Note.Split(new[] { ':' }, 2);
                        if (parts.Length > 0 && DateTime.TryParse(parts[0], out DateTime commentDate))
                        {
                            if (commentDate < cutoffDate)
                                remove = true;
                        }
                    }

                    if (remove)
                        comments.RemoveAt(i);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
