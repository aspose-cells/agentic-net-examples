// Title: Add a “Pending review” comment to every cell in the named range “ReviewNotes” using Aspose.Cells for .NET (C#)
// Description: This example loads an existing workbook, finds the named range "ReviewNotes", calculates its boundaries, iterates through each cell in the range, adds a comment with the text "Pending review", and saves the updated file. It demonstrates safe file handling, named‑range lookup, and comment creation with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | named range | add comment | Excel comment | ReviewNotes | cell iteration | Excel automation | pending review | batch comment
// Common Searches: Aspose.Cells add comment to all cells in a named range | C# add pending review comment to Excel range | How to loop through cells of a named range with Aspose.Cells | Add comment to Excel cells using Aspose.Cells .NET | Set comment for each cell in ReviewNotes range
// Developer Intent: Insert a "Pending review" comment into every cell that belongs to the named range "ReviewNotes" in an Excel workbook.
// Use Cases: Flag cells that require validation during a data‑quality audit. | Automatically annotate generated report sections so reviewers can see pending items. | Prepare a template with review notes before distributing it to stakeholders.
// AI Prompts: Generate C# code with Aspose.Cells that adds a custom comment to each cell of a specified named range and saves the workbook. | Show an alternative Aspose.Cells technique to apply the same comment to a named range without explicit loops. | Explain how to detect existing comments and update them only when the "Pending review" note is missing.

using System;
using System.IO;
using Aspose.Cells;

// This example loads an existing workbook, finds the named range "ReviewNotes", calculates its boundaries, iterates through each cell in the range, adds a comment with the text "Pending review", and saves the updated file. It demonstrates safe file handling, named‑range lookup, and comment creation with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Locate the named range "ReviewNotes"
            Name reviewName = null;
            foreach (Name n in workbook.Worksheets.Names)
            {
                if (n.Text == "ReviewNotes")
                {
                    reviewName = n;
                    break;
                }
            }

            if (reviewName == null)
            {
                Console.WriteLine("Named range 'ReviewNotes' not found.");
                return;
            }

            // Get the range object and its worksheet
            Aspose.Cells.Range reviewRange = reviewName.GetRange();
            Worksheet sheet = reviewRange.Worksheet;

            // Determine the boundaries of the range
            int startRow = reviewRange.FirstRow;
            int startCol = reviewRange.FirstColumn;
            int endRow = startRow + reviewRange.RowCount - 1;
            int endCol = startCol + reviewRange.ColumnCount - 1;

            // Add a comment to each cell in the range
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    int commentIndex = sheet.Comments.Add(row, col);
                    Comment comment = sheet.Comments[commentIndex];
                    comment.Note = "Pending review";
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
