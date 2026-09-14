// Title: Extract all worksheet comments from an Excel file and save them as indented JSON using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx workbook with Aspose.Cells, iterates through every worksheet's CommentCollection, captures each comment's author, cell address, and text, and writes the data to a pretty‑printed JSON file. | Create a reusable method that takes an input workbook path and an output JSON path, extracts comment details into a list of POCO objects, and serializes the list with System.Text.Json while ensuring the output directory exists.
// Common Searches: how to read cell comments from all sheets with Aspose.Cells in C# and export to JSON | C# Aspose.Cells extract worksheet notes and generate a formatted JSON report | save Excel comments to JSON file using Aspose.Cells .NET library | export workbook comment author and text to JSON with Aspose.Cells
// Tags: Aspose.Cells comment collection extraction | C# write Excel comments to JSON | System.Text.Json serialization of cell notes | export workbook comments as formatted JSON | iterate worksheets to gather comments Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsCommentExport
{
    // Model representing a comment extracted from a worksheet
    // The sample loads an Excel workbook via Aspose.Cells, loops through each worksheet's CommentCollection, records the worksheet name, cell address, author, and note into CellCommentInfo objects, aggregates them in a list, and then serializes the list to an indented JSON file, creating the output directory if it does not exist.
    public class CellCommentInfo
    {
        public string? WorksheetName { get; set; }
        public string? CellName { get; set; }
        public string? Author { get; set; }
        public string? Text { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Input Excel file path
            string excelPath = @"C:\Path\To\InputWorkbook.xlsx";

            // Output JSON file path
            string jsonPath = @"C:\Path\To\CommentsExport.json";

            try
            {
                // Verify input file exists
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"Input file not found: {excelPath}");
                    return;
                }

                // Load the workbook using Aspose.Cells
                Workbook workbook = new Workbook(excelPath);

                // List to hold all extracted comments
                List<CellCommentInfo> comments = new List<CellCommentInfo>();

                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Access the collection of comments for the current worksheet
                    CommentCollection commentCollection = sheet.Comments;

                    // Iterate through each comment in the collection
                    foreach (Comment comment in commentCollection)
                    {
                        // Build a comment info object
                        CellCommentInfo info = new CellCommentInfo
                        {
                            WorksheetName = sheet.Name,
                            CellName = CellsHelper.CellIndexToName(comment.Row, comment.Column),
                            Author = comment.Author,
                            Text = comment.Note
                        };

                        comments.Add(info);
                    }
                }

                // Serialize the list of comments to JSON with indentation for readability
                JsonSerializerOptions options = new JsonSerializerOptions
                {
                    WriteIndented = true
                };
                string json = JsonSerializer.Serialize(comments, options);

                // Ensure output directory exists
                string? outputDir = Path.GetDirectoryName(jsonPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Write the JSON string to the output file
                File.WriteAllText(jsonPath, json);

                Console.WriteLine($"Exported {comments.Count} comments to '{jsonPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
