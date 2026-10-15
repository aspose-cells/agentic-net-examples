// Title: Extract distinct authors of threaded comments in column B from an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that opens an existing .xlsx file, scans all worksheets, and returns a list of unique comment authors located in column B. | Show how to safely handle missing files and null author values while collecting comment authors from a specific column using Aspose.Cells.
// Common Searches: Aspose.Cells C# get unique comment authors from column B | How to list authors of Excel comments in a specific column using Aspose.Cells | Retrieve distinct comment author names across all worksheets with Aspose.Cells .NET | C# Aspose.Cells filter comments by column index | Extract comment authors from Excel file with Aspose.Cells handling null values
// Tags: Aspose.Cells extract comment authors | C# filter comments by column Aspose.Cells | retrieve distinct comment authors Excel .NET | handle null comment author Aspose.Cells | iterate worksheets comments Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// Loads input.xlsx, iterates through each worksheet, collects distinct authors of threaded comments located in column B (zero‑based index 1) into a HashSet, and prints each author while safely handling missing files and null author values.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            var authors = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the collection of (regular) comments on the worksheet
                CommentCollection comments = sheet.Comments;

                // Loop through each comment
                foreach (Comment comment in comments)
                {
                    // Column index is zero‑based; column B => index 1
                    if (comment.Column == 1)
                    {
                        // Add the author's name to the set (handle possible null)
                        authors.Add(comment.Author ?? "Unknown");
                    }
                }
            }

            // Output distinct authors
            foreach (string author in authors)
            {
                Console.WriteLine(author);
            }
        }
        catch (Exception ex)
        {
            // Runtime safety: report any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
