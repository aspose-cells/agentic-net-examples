// Title: Delete Excel worksheet comments older than 30 days using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that loops through every worksheet, parses each comment's note as a DateTime, and removes comments whose date is earlier than 30 days from today. | Show how to safely delete comments from a CommentCollection by iterating backwards, applying a cutoff date, and saving the workbook in XLSX format. | Provide an example that checks for a parsable date in a comment, compares it to a cutoff, and removes stale comments before saving.
// Common Searches: c# aspose.cells delete worksheet comments older than 30 days | how to remove Excel notes based on creation date using Aspose.Cells | filter and purge stale comments from an Excel file with Aspose.Cells .NET | aspnet remove old Excel comments programmatically with Aspose.Cells | parse comment note as date and delete outdated comments in Excel using Aspose.Cells
// Tags: Aspose.Cells remove comments by date | C# iterate worksheet CommentCollection | delete Excel notes older than 30 days | Aspose.Cells comment cleanup | filter Excel comments using parsed timestamp | save workbook after comment removal Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using System.Collections.Generic;

namespace AsposeCellsCommentCleanup
{
    // The sample loads an input XLSX file, iterates each worksheet's CommentCollection, parses a date from each comment's Note, removes comments older than a 30‑day cutoff, and saves the cleaned workbook as an output XLSX file.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Define the cutoff date: comments older than 30 days will be removed
                DateTime cutoffDate = DateTime.Now.AddDays(-30);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    CommentCollection comments = sheet.Comments;

                    // Iterate backwards to safely remove items
                    for (int i = comments.Count - 1; i >= 0; i--)
                    {
                        Comment comment = comments[i];

                        // Attempt to parse a date from the comment's note text (if present)
                        // Adjust the parsing logic according to your actual comment format
                        if (DateTime.TryParse(comment.Note, out DateTime commentDate))
                        {
                            if (commentDate < cutoffDate)
                            {
                                comments.RemoveAt(i);
                            }
                        }
                        else
                        {
                            // If no parsable date is found, optionally keep or remove the comment.
                            // Here we keep the comment.
                        }
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
}
