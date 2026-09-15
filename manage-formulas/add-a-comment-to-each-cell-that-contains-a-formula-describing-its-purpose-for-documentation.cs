// Title: Add visible documentation comments to every formula cell in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that scans all worksheets in a workbook and attaches a visible comment to each cell that contains a formula, where the comment text includes the exact formula string. | Update the sample program to prepend the cell address (e.g., A1) to the comment text and set the comment author from a variable named commentAuthor. | Create a C# version that extracts all formulas from the source workbook and writes them as comments on a new worksheet called "FormulaDocs" instead of attaching them to the original cells.
// Common Searches: how to automatically add comments to formula cells using Aspose.Cells in C# | Aspose.Cells C# iterate through all cells and document formulas with comments | add visible comment with formula text to each Excel cell using Aspose.Cells .NET | C# generate documentation for Excel formulas by inserting cell comments via Aspose.Cells | save workbook after adding comments to formula cells with Aspose.Cells
// Tags: Aspose.Cells insert formula comment | C# loop worksheets Aspose.Cells | Excel formula documentation via cell comments | make cell comments visible Aspose.Cells | save annotated workbook Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an input Excel file, iterates through every worksheet and each used cell, and for every cell that contains a formula it creates a visible comment whose note includes the formula text. The comment author is set to "Documentation" and the workbook is saved to the specified output path.
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
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                Cells cells = worksheet.Cells;

                // Loop over all used cells in the worksheet
                foreach (Cell cell in cells)
                {
                    // Process only cells that contain a formula
                    if (cell.IsFormula)
                    {
                        // Create a comment describing the formula
                        string commentText = $"Formula: {cell.Formula}";

                        // Add a new comment to the cell
                        int commentIndex = worksheet.Comments.Add(cell.Row, cell.Column);
                        Comment comment = worksheet.Comments[commentIndex];
                        comment.Author = "Documentation";
                        comment.Note = commentText;
                        // Optionally make the comment visible
                        comment.IsVisible = true;
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
