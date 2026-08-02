// Title: Export Excel worksheet comments to JSON using Aspose.Cells for .NET
// Description: Loads an Excel workbook with Aspose.Cells, iterates every worksheet's CommentCollection, converts each comment's row/column to A1 notation, captures sheet name, cell address, comment text and author, serializes the data to indented JSON, and writes it to a specified file while ensuring the output folder exists.
// Keywords: Aspose.Cells | export Excel comments to JSON | C# read cell comments | worksheet comment extraction | A1 address conversion | JSON serialization .NET | Excel comment author | Aspose.Cells .NET example
// Common Searches: How to extract Excel comments with Aspose.Cells C# | Export worksheet comments to JSON file | Convert comment row and column to A1 address in C# | Aspose.Cells get comment author and text | Save Excel cell comments as JSON using .NET
// Developer Intent: Retrieve all comments from an Excel workbook and store them as structured JSON.
// Use Cases: Create a JSON audit trail of documentation comments across all sheets. | Feed comment metadata (sheet, cell, author, text) into a reporting or analytics system. | Back up cell comments before performing bulk modifications or migrations.
// AI Prompts: Generate C# code that uses Aspose.Cells to read every comment in a workbook and output a formatted JSON array with sheet name, cell address, comment text, and author. | Show how to convert comment row and column indices to A1 notation while exporting comments to JSON. | Explain error handling for missing workbook files and automatic creation of the output directory when exporting comments.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsCommentExport
{
    // POCO to hold comment information for JSON serialization
    // Loads an Excel workbook with Aspose.Cells, iterates every worksheet's CommentCollection, converts each comment's row/column to A1 notation, captures sheet name, cell address, comment text and author, serializes the data to indented JSON, and writes it to a specified file while ensuring the output folder exists.
    public class CellCommentInfo
    {
        public string? SheetName { get; set; }   // Name of the worksheet
        public string? CellName { get; set; }    // Address of the cell (e.g., "A1")
        public string? Note { get; set; }        // Comment text
        public string? Author { get; set; }      // Comment author (if set)
    }

    public class ExportCommentsToJson
    {
        public static void Run(string workbookPath, string jsonOutputPath)
        {
            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(workbookPath))
                {
                    Console.Error.WriteLine($"Error: Workbook file not found at '{workbookPath}'.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Collect comment data from all worksheets
                List<CellCommentInfo> commentsData = new List<CellCommentInfo>();

                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    CommentCollection comments = sheet.Comments;

                    for (int i = 0; i < comments.Count; i++)
                    {
                        Comment comment = comments[i];

                        // Convert row/column indices to an A1 style address
                        string cellAddress = CellsHelper.CellIndexToName(comment.Row, comment.Column);

                        commentsData.Add(new CellCommentInfo
                        {
                            SheetName = sheet.Name,
                            CellName = cellAddress,
                            Note = comment.Note,
                            Author = comment.Author
                        });
                    }
                }

                // Serialize to JSON with indentation
                JsonSerializerOptions jsonOptions = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(commentsData, jsonOptions);

                // Ensure the output directory exists
                string? outputDir = Path.GetDirectoryName(jsonOutputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Write JSON to file
                File.WriteAllText(jsonOutputPath, json);

                Console.WriteLine($"Exported {commentsData.Count} comments to '{jsonOutputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the existing Excel workbook
            string inputPath = "InputWorkbook.xlsx";

            // Desired path for the JSON output
            string outputPath = "CommentsExport.json";

            // Execute the export
            ExportCommentsToJson.Run(inputPath, outputPath);
        }
    }
}
