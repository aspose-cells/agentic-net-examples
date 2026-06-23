using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsCommentExport
{
    // Simple POCO to hold comment information for JSON serialization
    public class CellCommentInfo
    {
        public string Cell { get; set; }
        public string Note { get; set; }
        public string Author { get; set; }
    }

    public class ExportCommentsToJson
    {
        public static void Run()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Prepare a list to collect comment data
            List<CellCommentInfo> commentsData = new List<CellCommentInfo>();

            // Iterate through all comments in the worksheet
            foreach (Comment comment in worksheet.Comments)
            {
                // Retrieve cell address (e.g., "A1")
                string cellName = worksheet.Cells[comment.Row, comment.Column].Name;

                // Build the comment info object
                CellCommentInfo info = new CellCommentInfo
                {
                    Cell = cellName,
                    Note = comment.Note,
                    Author = comment.Author
                };

                commentsData.Add(info);
            }

            // Serialize the list to a formatted JSON string
            string jsonOutput = JsonSerializer.Serialize(commentsData, new JsonSerializerOptions
            {
                WriteIndented = true
            });

            // Write JSON to a file
            File.WriteAllText("comments.json", jsonOutput);

            Console.WriteLine("Comments have been exported to comments.json");
        }
    }

    // Entry point for demonstration
    class Program
    {
        static void Main(string[] args)
        {
            ExportCommentsToJson.Run();
        }
    }
}