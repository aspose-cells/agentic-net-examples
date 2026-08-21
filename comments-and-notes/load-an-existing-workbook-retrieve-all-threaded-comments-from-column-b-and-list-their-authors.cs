// Title: C# – List unique threaded comment authors in column B using Aspose.Cells
// Description: Loads an Excel workbook with Aspose.Cells, scans all used rows in column B, extracts any threaded comments, gathers distinct author names, and prints the list to the console.
// Keywords: Aspose.Cells | C# | threaded comments | Excel author extraction | column B | unique comment authors | GetThreadedComments | Workbook | Worksheet | HashSet
// Common Searches: Aspose.Cells get threaded comment authors C# | list distinct comment authors column B Excel | retrieve authors from Excel comments using .NET | how to extract threaded comment authors Aspose.Cells | C# code to read Excel comment authors
// Developer Intent: Extract and display each unique author of threaded comments located in column B of an Excel worksheet.
// Use Cases: Create an audit log of users who have left threaded comments in column B before releasing the workbook. | Validate that only approved personnel have commented in column B by comparing extracted names to a whitelist. | Populate a dropdown or filter UI with distinct comment authors from column B for further analysis.
// AI Prompts: Generate C# code with Aspose.Cells that collects distinct author names from threaded comments in column B of an Excel file. | Show how to export the unique list of comment authors from column B to a CSV using Aspose.Cells. | Explain best practices for handling cells without threaded comments when extracting authors in a .NET application.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Loads an Excel workbook with Aspose.Cells, scans all used rows in column B, extracts any threaded comments, gathers distinct author names, and prints the list to the console.
class ThreadedCommentAuthorsFromColumnB
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("InputWorkbook.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Collection to store unique author names
        HashSet<string> authors = new HashSet<string>();

        // Determine the last used row in the worksheet
        int lastRow = worksheet.Cells.MaxDataRow;

        // Column B has index 1 (zero‑based)
        int columnIndex = 1;

        // Iterate through each row in column B
        for (int row = 0; row <= lastRow; row++)
        {
            // Retrieve threaded comments for the current cell (row, column B)
            ThreadedCommentCollection threadedComments = worksheet.Comments.GetThreadedComments(row, columnIndex);

            // If there are no threaded comments, continue to next row
            if (threadedComments == null || threadedComments.Count == 0)
                continue;

            // Extract author names from each threaded comment
            foreach (ThreadedComment comment in threadedComments)
            {
                if (comment?.Author != null && !string.IsNullOrEmpty(comment.Author.Name))
                {
                    authors.Add(comment.Author.Name);
                }
            }
        }

        // List all distinct authors found in column B
        Console.WriteLine("Threaded comment authors in column B:");
        foreach (string authorName in authors)
        {
            Console.WriteLine("- " + authorName);
        }
    }
}
