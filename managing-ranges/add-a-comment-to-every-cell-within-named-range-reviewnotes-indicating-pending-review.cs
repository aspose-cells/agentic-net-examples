// Title: Add a “Pending review” comment to each cell in the named range “ReviewNotes” using Aspose.Cells for .NET (C#)
// Description: C# example that loads an Excel workbook with Aspose.Cells, checks for the named range "ReviewNotes", iterates through every cell in that range, adds a "Pending review" comment, and saves the updated file. Includes file‑existence and range‑validation checks.
// Keywords: Aspose.Cells comment API | C# add comment to Excel cell | named range comment Aspose | ReviewNotes Excel range | iterate cells Aspose.Cells | pending review tag Excel | programmatic Excel comment | Excel workflow tagging
// Common Searches: Aspose.Cells add comment to named range C# | C# add comment to every cell in Excel range | How to tag cells with "Pending review" using Aspose.Cells | Iterate over a named range and insert comments in .NET | Add comment to ReviewNotes range programmatically
// Developer Intent: Insert a "Pending review" comment into every cell that belongs to the named range "ReviewNotes" in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Automatically flag cells that require validation before publishing a report. | Create a review workflow by marking all cells in a specific range with a pending‑review comment. | Maintain an audit trail by programmatically adding comments to cells defined by a named range. | Prepare spreadsheets for collaborative editing where reviewers can see which cells need attention.
// AI Prompts: Generate C# code with Aspose.Cells that adds a custom comment to each cell in a named range called "ReviewNotes" and saves the workbook. | Explain how to safely verify the existence of a named range before adding comments using Aspose.Cells for .NET. | Provide error‑handling best practices for missing input files and absent named ranges when inserting comments into an Excel file. | Show how to unit‑test the comment‑insertion logic for a named range with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Example class that adds a comment to each cell in the named range "ReviewNotes"
    // C# example that loads an Excel workbook with Aspose.Cells, checks for the named range "ReviewNotes", iterates through every cell in that range, adds a "Pending review" comment, and saves the updated file. Includes file‑existence and range‑validation checks.
    public class AddCommentToReviewNotesRange
    {
        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Ensure the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Locate the named range "ReviewNotes"
                Name reviewName = workbook.Worksheets.Names["ReviewNotes"];
                if (reviewName == null)
                {
                    Console.WriteLine("Named range 'ReviewNotes' not found.");
                    return;
                }

                // Get the range that the name refers to
                Aspose.Cells.Range reviewRange = reviewName.GetRange();

                // Iterate through each cell in the range and add a comment
                int startRow = reviewRange.FirstRow;
                int startColumn = reviewRange.FirstColumn;
                int rowCount = reviewRange.RowCount;
                int columnCount = reviewRange.ColumnCount;

                for (int i = 0; i < rowCount; i++)
                {
                    for (int j = 0; j < columnCount; j++)
                    {
                        int row = startRow + i;
                        int col = startColumn + j;

                        // Add a comment to the current cell
                        int commentIndex = worksheet.Comments.Add(row, col);
                        Comment comment = worksheet.Comments[commentIndex];
                        comment.Note = "Pending review";
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            AddCommentToReviewNotesRange.Run();
        }
    }
}
