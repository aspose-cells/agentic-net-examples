using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

class ExportCommentsToJson
{
    static void Main()
    {
        // Paths
        string inputPath = "input.xlsx";
        string jsonPath = "comments.json";

        // Verify input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load workbook
            Workbook workbook = new Workbook(inputPath);
            var comments = new List<object>();

            // Iterate worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate all cells in the worksheet
                foreach (Cell cell in sheet.Cells)
                {
                    // If the cell has a comment, capture its details
                    if (cell?.Comment != null)
                    {
                        comments.Add(new
                        {
                            Sheet = sheet.Name,
                            Cell = cell.Name,
                            Author = cell.Comment.Author,
                            Note = cell.Comment.Note
                        });
                    }
                }
            }

            // Serialize to JSON
            var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
            string json = JsonSerializer.Serialize(comments, jsonOptions);

            // Write JSON file (ensure directory exists)
            try
            {
                File.WriteAllText(jsonPath, json);
                Console.WriteLine($"Comments exported to {jsonPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write JSON file: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
        }
    }
}