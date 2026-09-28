// Title: How to edit the text of an existing cell comment in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Update the Note property of the comment attached to cell B2 and save the workbook as a new file with Aspose.Cells. | Replace the content of a specific worksheet comment in C# without creating a new comment object using the Aspose.Cells API.
// Common Searches: aspocells change comment text in a specific cell c# example | c# update existing comment note in excel workbook using Aspose.Cells | modify worksheet comment programmatically aspocells .net | save workbook after editing cell comment aspocells c#
// Tags: Aspose.Cells edit comment Note property | C# update worksheet comment text | Aspose.Cells save workbook after comment change | Excel comment editing with Aspose.Cells .NET | cell B2 comment update Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an existing Excel file, accesses cell B2 on the first worksheet, checks for a comment, updates its Note to a new string, ensures the output directory exists, and saves the modified workbook to a new file.
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (or specify by name)
            Worksheet sheet = workbook.Worksheets[0];

            // Resolve cell B2 to row/column indices
            Cell targetCell = sheet.Cells["B2"];
            int row = targetCell.Row;
            int column = targetCell.Column;

            // Retrieve a regular comment for cell B2 using the Comments collection
            Comment comment = sheet.Comments[row, column];

            if (comment != null)
            {
                // Edit the comment text
                comment.Note = "This is the updated comment text.";
            }
            else
            {
                Console.WriteLine("No comment found in cell B2.");
            }

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the modified comment (lifecycle rule: save)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
