// Title: Change an Excel comment’s font color to red using Aspose.Cells for .NET (C#)
// Common Searches: how to set the font color of an Excel comment using Aspose.Cells in C# | Aspose.Cells C# change comment text color to red | programmatically modify comment shape font color in a .NET workbook | C# code example for changing Excel comment font color with Aspose.Cells | set comment shape Font.Color property Aspose.Cells example
// Tags: Aspose.Cells comment font color | C# Excel comment shape formatting | set comment text color Aspose.Cells | change comment font color .NET | Excel comment shape Font.Color property | modify comment appearance using Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an existing workbook, ensures a comment exists on cell B2, accesses the comment's shape, sets its Font.Color to red, and saves the updated file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Target cell for the comment (e.g., B2)
            Cell targetCell = worksheet.Cells["B2"];

            // Retrieve existing comment or add a new one
            Comment comment = worksheet.Comments[targetCell.Name];
            if (comment == null)
            {
                // Add a new comment to the cell
                int commentIndex = worksheet.Comments.Add(targetCell.Name);
                comment = worksheet.Comments[commentIndex];
                comment.Note = "This is a sample comment.";
            }

            // Access the shape that represents the comment
            CommentShape commentShape = comment.CommentShape;

            // Change the font color of the comment text to red
            commentShape.Font.Color = Color.Red;

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
