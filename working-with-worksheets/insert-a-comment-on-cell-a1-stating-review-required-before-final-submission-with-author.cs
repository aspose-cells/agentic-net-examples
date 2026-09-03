// Title: Insert a comment with author into cell A1 of a new Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a fresh Workbook, accesses the first worksheet, adds a comment to cell A1, sets the comment's Author to "AuthorName" and its Note to "Review required before final submission", ensures the output directory exists, and saves the file as Output.xlsx with Aspose.Cells. | Generate a .NET example that demonstrates how to place a reviewer note in the first cell, assign an author name to the comment, and export the workbook to a specified path while handling missing folders, using Aspose.Cells.
// Common Searches: asp.net insert reviewer note into the first cell with Aspose.Cells | c# Aspose.Cells set comment author and text for a specific worksheet cell | how to ensure output directory exists before saving Excel file with Aspose.Cells | example of saving a workbook as Output.xlsx after adding a comment in C#
// Tags: add comment to worksheet cell Aspose.Cells C# | assign author to Excel comment Aspose.Cells | create workbook and insert cell note Aspose.Cells | save workbook with folder creation Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The program creates a new Workbook, accesses the first worksheet, inserts a comment with author "AuthorName" and note "Review required before final submission" into cell A1, ensures the target directory exists, and saves the workbook as Output.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add a comment to cell A1 (row 0, column 0)
                // The indexer creates the comment if it does not exist
                Comment comment = sheet.Comments[0, 0];
                comment.Author = "AuthorName";
                comment.Note = "Review required before final submission";

                // Define output file path
                string outputPath = "Output.xlsx";

                // Ensure the output directory exists (if a directory part is present)
                string? outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
