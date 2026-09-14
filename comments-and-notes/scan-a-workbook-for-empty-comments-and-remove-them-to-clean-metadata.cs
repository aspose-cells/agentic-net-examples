// Title: Remove empty cell comments from all worksheets in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that loads an XLSX file, scans every worksheet, and deletes comments whose Note property is null, empty, or whitespace, then saves the cleaned file. | Add detailed console logging to the RemoveEmptyComments example to output the address of each comment that is removed. | Create a reusable static method DeleteBlankComments(string sourcePath, string destinationPath) that uses Aspose.Cells to purge any Excel workbook of empty comments.
// Common Searches: c# aspocells how to delete comments with no text from an Excel file | remove blank notes from all sheets using Aspose.Cells for .NET | Aspose.Cells iterate through worksheet comments and filter out empty ones | clean up Excel comment metadata programmatically with Aspose.Cells C# | delete whitespace-only comments in XLSX using Aspose.Cells library
// Tags: Aspose.Cells delete empty comments | C# remove blank worksheet notes | clean Excel comment metadata Aspose.Cells | iterate worksheet comments C# | remove whitespace comments XLSX

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example loads an Excel workbook, iterates each worksheet to locate comments whose Note property is null, empty, or whitespace, removes those empty comments, and saves the cleaned workbook as an XLSX file.
    public class RemoveEmptyComments
    {
        public static void Run(string inputPath, string outputPath)
        {
            try
            {
                // Load the workbook (input file existence already verified)
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    CommentCollection comments = worksheet.Comments;
                    var emptyCommentPositions = new List<(int Row, int Column)>();

                    // Identify empty comments
                    foreach (Comment comment in comments)
                    {
                        if (string.IsNullOrWhiteSpace(comment.Note))
                        {
                            emptyCommentPositions.Add((comment.Row, comment.Column));
                        }
                    }

                    // Remove empty comments
                    foreach ((int row, int column) in emptyCommentPositions)
                    {
                        comments.RemoveAt(row, column);
                    }
                }

                // Save the cleaned workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Ensure the input file exists; create a placeholder if missing
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found. Creating an empty workbook as placeholder.");
                    Workbook placeholder = new Workbook();
                    placeholder.Save(inputPath, SaveFormat.Xlsx);
                }

                Run(inputPath, outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
