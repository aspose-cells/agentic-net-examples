// Title: C# – Load an Excel workbook and list distinct authors of threaded comments in column B using Aspose.Cells
// Description: Shows how to open an existing .xlsx file with Aspose.Cells for .NET, iterate through all occupied rows in column B, retrieve threaded comments from each cell, print the cell address with the comment author, and finally output a unique list of authors.
// Keywords: Aspose.Cells | C# | .NET | Excel | threaded comments | comment authors | column B | list distinct authors | retrieve comments | Workbook | Worksheet | GetThreadedComments | author extraction | Excel automation | Aspose.Cells example
// Common Searches: Aspose.Cells get threaded comment authors column B | C# list distinct comment authors in Excel using Aspose.Cells | How to retrieve threaded comments from a specific column with Aspose.Cells | Extract unique comment authors from an Excel worksheet .NET | Aspose.Cells example for reading threaded comments
// Developer Intent: Extract and display unique authors of threaded comments in column B of an Excel workbook.
// Use Cases: Generate a report of users who have commented in a specific column. | Validate comment ownership before further data processing. | Create an audit trail of comment activity per column. | Integrate author extraction into a larger data‑analysis pipeline.
// AI Prompts: Write a C# method using Aspose.Cells that returns a HashSet<string> of distinct threaded comment authors from a given column index. | Explain how to safely handle empty worksheets when retrieving threaded comments with Aspose.Cells. | Provide a complete console application that prints each cell address and its comment author for column C. | Show how to filter threaded comments by author name using Aspose.Cells.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Shows how to open an existing .xlsx file with Aspose.Cells for .NET, iterate through all occupied rows in column B, retrieve threaded comments from each cell, print the cell address with the comment author, and finally output a unique list of authors.
class Program
{
    static void Main()
    {
        // Load the existing workbook (replace with your file path)
        Workbook workbook = new Workbook("Input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Column B has index 1 (zero‑based)
        int columnIndex = 1;

        // Determine the last row that contains data; if the sheet is empty, start at row 0
        int maxRow = worksheet.Cells.MaxDataRow;
        if (maxRow < 0) maxRow = 0;

        // Use a set to collect distinct author names
        HashSet<string> distinctAuthors = new HashSet<string>();

        // Iterate through each row in column B
        for (int row = 0; row <= maxRow; row++)
        {
            // Retrieve threaded comments for the current cell (row, column B)
            ThreadedCommentCollection threadedComments = worksheet.Comments.GetThreadedComments(row, columnIndex);

            // If there are any threaded comments, process them
            if (threadedComments != null && threadedComments.Count > 0)
            {
                foreach (ThreadedComment comment in threadedComments)
                {
                    if (comment.Author != null)
                    {
                        string authorName = comment.Author.Name;
                        distinctAuthors.Add(authorName);

                        // Output the cell address and its comment author
                        Console.WriteLine($"Cell {CellsHelper.CellIndexToName(row, columnIndex)} - Author: {authorName}");
                    }
                }
            }
        }

        // List all distinct authors found in column B
        Console.WriteLine("\nDistinct authors in column B:");
        foreach (string name in distinctAuthors)
        {
            Console.WriteLine(name);
        }
    }
}
