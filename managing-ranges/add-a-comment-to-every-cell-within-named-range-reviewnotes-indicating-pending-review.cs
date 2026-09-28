// Title: Add a “Pending review” comment to every cell in the named range “ReviewNotes” using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, retrieves the named range "ReviewNotes", and adds a "Pending review" comment to each cell, creating the comment if it does not exist. | Adjust the Aspose.Cells loop to set the comment author from a variable and ensure the workbook is saved after all comments are updated. | Create error‑handling logic that logs the cell address when adding or updating a comment fails while processing a named range.
// Common Searches: Aspose.Cells C# add comment to all cells in a named range | How to set comment text for each cell in a specific range using Aspose.Cells | C# update existing cell comments in an Excel workbook with Aspose.Cells | Iterate over named range ReviewNotes and add pending review comment Aspose | Save workbook after modifying comments with Aspose.Cells .NET
// Tags: Aspose.Cells add comment to range | update Excel cell comments programmatically | named range comment handling Aspose | C# iterate Aspose.Range cells | save workbook after comment changes Aspose

using Aspose.Cells;
using System;
using System.IO;

// Alias to avoid conflict with System.Range introduced in newer C# versions
using AsposeRange = Aspose.Cells.Range;

// // Loads "input.xlsx", finds the named range "ReviewNotes", adds or updates a "Pending review" comment for every cell in that range, and saves the modified workbook as "output.xlsx".
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Retrieve the named range "ReviewNotes"
            var namedRange = workbook.Worksheets.Names["ReviewNotes"];
            if (namedRange != null)
            {
                // Obtain the actual cell range represented by the named range
                AsposeRange range = namedRange.GetRange();

                // Iterate over each cell in the range
                for (int row = range.FirstRow; row < range.FirstRow + range.RowCount; row++)
                {
                    for (int col = range.FirstColumn; col < range.FirstColumn + range.ColumnCount; col++)
                    {
                        Cell cell = range.Worksheet.Cells[row, col];

                        try
                        {
                            // Add or update a comment indicating pending review
                            var comment = cell.Comment;
                            if (comment == null)
                            {
                                // Create a new comment
                                int commentIndex = range.Worksheet.Comments.Add(row, col);
                                comment = range.Worksheet.Comments[commentIndex];
                                comment.Author = "Author";
                                comment.Note = "Pending review";
                            }
                            else
                            {
                                // Update existing comment
                                comment.Note = "Pending review";
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Failed to process comment for cell {cell.Name}: {ex.Message}");
                        }
                    }
                }
            }
            else
            {
                Console.WriteLine("Named range 'ReviewNotes' not found.");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
