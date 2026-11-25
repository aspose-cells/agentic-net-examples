// Title: Read all cell comments from an Excel workbook and save them as JSON using Aspose.Cells for .NET (C#)
// Common Searches: how to read Excel cell comments with Aspose.Cells in C# and output JSON | Aspose.Cells extract worksheet comment author and text to a JSON file | C# code to export Excel comments to a JSON array using Aspose.Cells | save Excel comments as JSON using the Aspose.Cells .NET library | retrieve comment visibility flag from an Excel sheet with Aspose.Cells and serialize it
// Tags: Aspose.Cells export comments to JSON | read worksheet comments C# | serialize Excel cell notes .NET | extract comment author Aspose.Cells | comment visibility Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace ExportCommentsToJson
{
    // Simple DTO to hold comment information for JSON serialization
    // The example loads an Excel workbook with Aspose.Cells, iterates through every worksheet's Comments collection, captures each comment's worksheet name, zero‑based row and column indices, author, note text, and visibility flag into a DTO, serializes the list to indented JSON, and writes the output to a user‑specified file while handling missing files and runtime exceptions.
    public class CellCommentInfo
    {
        public string Worksheet { get; set; }
        public int Row { get; set; }          // Zero‑based row index
        public int Column { get; set; }       // Zero‑based column index
        public string Author { get; set; }
        public string Note { get; set; }
        public bool IsVisible { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Determine input and output paths
            string inputPath = args.Length > 0 ? args[0] : "input.xlsx";
            string outputPath = args.Length > 1 ? args[1] : "comments.json";

            if (!File.Exists(inputPath))
            {
                Console.Error.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Collect all comments from all worksheets
                List<CellCommentInfo> allComments = new List<CellCommentInfo>();

                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // The Comments collection contains all comments on the sheet
                    foreach (Comment comment in sheet.Comments)
                    {
                        // Some versions expose Row/Column via comment.Row and comment.Column.
                        // If not available, we can retrieve the cell address from comment.Note's cell.
                        // Here we use the properties that are part of the SDK.
                        int row = comment.Row;
                        int column = comment.Column;

                        CellCommentInfo info = new CellCommentInfo
                        {
                            Worksheet = sheet.Name,
                            Row = row,
                            Column = column,
                            Author = comment.Author,
                            Note = comment.Note,
                            IsVisible = comment.IsVisible
                        };

                        allComments.Add(info);
                    }
                }

                // Serialize to JSON with indentation for readability
                var jsonOptions = new JsonSerializerOptions
                {
                    WriteIndented = true
                };
                string json = JsonSerializer.Serialize(allComments, jsonOptions);

                // Write JSON to file
                File.WriteAllText(outputPath, json);
                Console.WriteLine($"Successfully exported {allComments.Count} comment(s) to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
