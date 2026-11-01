// Title: C# example to delete empty cell comments (notes) from an Excel workbook using Aspose.Cells
// Common Searches: how to programmatically delete blank cell comments in an Excel file with Aspose.Cells C# | remove empty notes from all worksheets using Aspose.Cells for .NET | C# code to clean up Excel workbook by removing comments with no text | Aspose.Cells iterate through cells and delete comments that contain only whitespace | delete null or whitespace comments from Excel using Aspose.Cells API
// Tags: Aspose.Cells delete empty comments C# | remove blank notes Excel workbook .NET | iterate worksheets comment collection Aspose.Cells | clean Excel metadata by removing null comments | C# Excel comment cleanup Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The program loads an XLSX workbook with Aspose.Cells, iterates over every worksheet and each used cell, checks the cell's comment (note) and removes it when the comment text is null, empty, or whitespace, then saves the cleaned workbook to a new file.
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
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    Cells cells = sheet.Cells;

                    // Iterate through all used cells in the worksheet
                    foreach (Cell cell in cells)
                    {
                        // Retrieve the comment (if any) from the worksheet's comment collection
                        Comment comment = sheet.Comments[cell.Row, cell.Column];
                        if (comment != null)
                        {
                            // Get the comment text (Aspose.Cells uses the Note property)
                            string commentText = comment.Note;

                            // If the comment text is null, empty, or whitespace, remove the comment
                            if (string.IsNullOrWhiteSpace(commentText))
                            {
                                sheet.Comments.RemoveAt(cell.Row, cell.Column);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing sheet \"{sheet.Name}\": {ex.Message}");
                }
            }

            // Save the cleaned workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
