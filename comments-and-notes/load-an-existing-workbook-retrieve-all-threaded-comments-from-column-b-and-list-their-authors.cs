// Title: Load an Excel workbook and list comment authors from column B using Aspose.Cells for .NET (C#)
// Common Searches: asp.net c# read comment author from column B in existing Excel file using Aspose.Cells | how to filter worksheet comments by column index with Aspose.Cells in C# | retrieve threaded comment authors from a specific column in an Excel workbook with Aspose.Cells | c# iterate Excel comments and output authors for column B using Aspose.Cells library | list all comment authors in column B of a loaded workbook with Aspose.Cells for .NET
// Tags: Aspose.Cells load workbook | Aspose.Cells read worksheet comments | Aspose.Cells filter comments by column | Aspose.Cells get comment author | C# extract Excel comment authors | Excel column B comment retrieval Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks for 'input.xlsx', loads it with Aspose.Cells, accesses the first worksheet, iterates over all worksheet comments, filters those whose Column index equals 1 (column B), and writes each comment's Author to the console while handling potential exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Work with the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Iterate through all regular comments in the worksheet (fallback if threaded comments are unavailable)
            foreach (Comment comment in worksheet.Comments)
            {
                // Column B corresponds to index 1 (zero‑based)
                if (comment.Column == 1)
                {
                    // Output the author of the comment
                    Console.WriteLine(comment.Author);
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
