// Title: Extract and list comment authors from an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: how to read comment author names from an Excel file with Aspose.Cells in C# | C# code to get all comment authors from a worksheet using Aspose.Cells | retrieve list of cell comment authors from .xlsx using Aspose.Cells .NET library | Aspose.Cells iterate worksheet comments and collect author property C# | example of extracting comment authors from Excel workbook with Aspose.Cells
// Tags: Aspose.Cells retrieve comment author strings | C# iterate worksheet comment collection | list Excel comment authors .NET | handle null comment author Aspose.Cells | read comment metadata from .xlsx using Aspose

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The sample checks for an input.xlsx file, loads it with Aspose.Cells, accesses the first worksheet, iterates over its Comments collection, captures each comment's Author (defaulting to "Unknown" when null), stores the names in a list, and prints the author list to the console while handling any runtime exceptions.
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

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (or any specific worksheet as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Collect author names from all comments in the worksheet
            List<string> authors = new List<string>();

            // Iterate through the regular comments collection
            foreach (Comment comment in worksheet.Comments)
            {
                // Comment.Author returns the author's name as a string
                string authorName = comment.Author ?? "Unknown";
                authors.Add(authorName);
            }

            // Output the list of authors
            Console.WriteLine("Comment authors:");
            foreach (string author in authors)
            {
                Console.WriteLine(author);
            }
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
