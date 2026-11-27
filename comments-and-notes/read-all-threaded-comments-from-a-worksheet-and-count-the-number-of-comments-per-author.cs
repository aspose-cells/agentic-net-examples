// Title: C# code to count comments per author in an Excel worksheet using Aspose.Cells
// Common Searches: Aspose.Cells C# count how many comments each author has in an Excel sheet | Get comment author statistics from a .xlsx file with Aspose.Cells for .NET | Iterate worksheet comments and group by author using Aspose.Cells API | C# example to tally Excel cell comments per author with Aspose.Cells | How to handle missing comment author values when counting comments in Aspose.Cells
// Tags: Aspose.Cells comment author counting | enumerate worksheet comments .NET | group Excel comments by author C# | case‑insensitive dictionary for comment aggregation | load workbook and process comments Aspose.Cells | error handling for missing Excel file Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The sample loads an .xlsx workbook, selects the first worksheet, iterates through its Comments collection, and uses a case‑insensitive dictionary to tally the number of comments each author made. It prints each author with their comment count and includes checks for missing files and load errors.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: The file \"{inputFile}\" was not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputFile);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Choose the first worksheet for processing
        Worksheet worksheet = workbook.Worksheets[0];

        // Dictionary to store comment count per author (case‑insensitive)
        var authorCommentCount = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        try
        {
            // Iterate through all regular comments in the worksheet
            foreach (Comment comment in worksheet.Comments)
            {
                string author = comment.Author ?? "Unknown";

                if (authorCommentCount.ContainsKey(author))
                    authorCommentCount[author]++;
                else
                    authorCommentCount[author] = 1;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing comments: {ex.Message}");
            return;
        }

        // Output the results
        foreach (KeyValuePair<string, int> kvp in authorCommentCount)
        {
            Console.WriteLine($"Author: {kvp.Key}, Comments: {kvp.Value}");
        }
    }
}
