// Title: Remove a threaded comment from cell C3 in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# remove comment from cell C3 | Delete threaded comment at C3 with Worksheet.Comments.RemoveAt | Remove Excel cell comment programmatically using Aspose.Cells .NET | How to check and delete a comment in a specific cell using Aspose.Cells | C# code to remove a comment from an existing .xlsx file with Aspose
// Tags: Aspose.Cells threaded comment removal | Worksheet.Comments.RemoveAt usage | C# delete Excel cell comment | Aspose.Cells comment deletion workflow | save modified workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace RemoveThreadedCommentApp
{
    // The example loads an existing .xlsx file, accesses the first worksheet, checks for a threaded comment at cell C3 (row 3, column C), removes it using Worksheet.Comments.RemoveAt, ensures the output directory exists, and saves the workbook to a new file while handling missing files and runtime errors.
    class RemoveThreadedComment
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Ensure the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook from the input file
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Zero‑based indexes for cell C3
                int rowIndex = 2;      // Row 3
                int columnIndex = 2;   // Column C

                // Retrieve the comment at C3, if any
                Comment comment = worksheet.Comments[rowIndex, columnIndex];

                // Remove the comment using the worksheet's Comments collection
                if (comment != null)
                {
                    // Correct API to remove a comment by its position
                    worksheet.Comments.RemoveAt(rowIndex, columnIndex);
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
