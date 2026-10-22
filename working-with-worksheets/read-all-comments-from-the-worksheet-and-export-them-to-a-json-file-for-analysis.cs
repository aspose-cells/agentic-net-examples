// Title: Export all Excel worksheet comments to a formatted JSON file with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an .xlsx workbook using Aspose.Cells, iterates through every worksheet, extracts each comment's author, note, and cell address, and writes the data to an indented JSON file. | Create a method that converts comment row and column indices to Excel cell names via CellsHelper.CellIndexToName and builds a List of objects containing WorksheetName, CellName, Author, and Note for JSON serialization. | Add robust error handling that verifies the input file exists, catches exceptions during workbook loading or JSON writing, and logs the total number of exported comments.
// Common Searches: how to extract comments from all sheets in an Excel file using Aspose.Cells C# | Aspose.Cells C# export cell comments to JSON file | read Excel comment author and note with Aspose.Cells in .NET | serialize Aspose.Cells comment collection to indented JSON | convert comment row column to cell address Aspose.Cells C#
// Tags: export worksheet comments to JSON Aspose.Cells | extract comment author and note Aspose.Cells | convert comment indices to cell name Aspose.Cells | serialize comment list to indented JSON C# | Aspose.Cells comment extraction sample

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsCommentExport
{
    // Model to hold comment information; properties are nullable to satisfy non‑nullable warnings
    // The program loads an Excel workbook, iterates through each worksheet, gathers every comment's worksheet name, cell address, author, and note, and writes the collected information as formatted JSON to an output file.
    public class CellCommentInfo
    {
        public string? WorksheetName { get; set; }
        public string? CellName { get; set; }
        public string? Author { get; set; }
        public string? Note { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "comments.json";

            try
            {
                // Ensure the input file exists before attempting to load it
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Collect comments from all worksheets
                List<CellCommentInfo> allComments = new List<CellCommentInfo>();

                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    CommentCollection comments = sheet.Comments;

                    foreach (Comment comment in comments)
                    {
                        // Convert row/column indices to Excel cell name (e.g., "B2")
                        string cellName = CellsHelper.CellIndexToName(comment.Row, comment.Column);

                        CellCommentInfo info = new CellCommentInfo
                        {
                            WorksheetName = sheet.Name,
                            CellName = cellName,
                            Author = comment.Author,
                            Note = comment.Note
                        };

                        allComments.Add(info);
                    }
                }

                // Serialize comments to formatted JSON
                JsonSerializerOptions options = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(allComments, options);

                // Write JSON to output file
                File.WriteAllText(outputPath, json);

                Console.WriteLine($"Exported {allComments.Count} comments to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
