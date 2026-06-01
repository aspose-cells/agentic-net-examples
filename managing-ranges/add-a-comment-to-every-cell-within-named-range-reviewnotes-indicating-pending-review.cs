using System;
using System.IO;
using Aspose.Cells;

class AddCommentsToNamedRange
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range "ReviewNotes"
            Name namedRange = workbook.Worksheets.Names["ReviewNotes"];
            if (namedRange == null)
            {
                Console.WriteLine("Named range 'ReviewNotes' not found.");
                return;
            }

            // Get the Range object that the name refers to (use fully qualified name to avoid ambiguity)
            Aspose.Cells.Range range = namedRange.GetRange();

            // Obtain the worksheet that contains the range
            Worksheet sheet = workbook.Worksheets[namedRange.SheetIndex];

            // Iterate through each cell in the range and add a comment
            for (int row = range.FirstRow; row < range.FirstRow + range.RowCount; row++)
            {
                for (int col = range.FirstColumn; col < range.FirstColumn + range.ColumnCount; col++)
                {
                    // Add a comment to the current cell
                    int commentIndex = sheet.Comments.Add(row, col);
                    Comment comment = sheet.Comments[commentIndex];
                    comment.Note = "Pending review";
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}