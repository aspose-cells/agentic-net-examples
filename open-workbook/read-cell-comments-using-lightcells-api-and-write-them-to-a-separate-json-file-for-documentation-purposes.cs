// Title: Read Excel cell comments with Aspose.Cells LightCells API and export them to a formatted JSON file using C#
// AI Prompts: Generate C# code that opens an .xlsx file with Aspose.Cells LightCells, iterates all worksheets, extracts each cell's comment text, and writes the results to an indented JSON document. | Modify the sample to include the comment author and timestamp in the JSON output while preserving the existing sheet name and address fields. | Create a reusable C# method that accepts a workbook path and returns a JSON string containing comment details (sheet, address, text, author, date) extracted via LightCells.
// Common Searches: how to extract cell comments from an Excel file using Aspose.Cells LightCells in C# | C# Aspose.Cells read comments and save as JSON for documentation | export Excel worksheet comments to JSON with LightCells API | Aspose.Cells LightCells iterate cells and get comment note in C# | serialize Excel cell notes to JSON file using Aspose.Cells
// Tags: Aspose.Cells LightCells comment extraction | convert Excel comments to JSON C# | read cell notes Aspose.Cells | serialize worksheet comments | C# workbook comment documentation

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// Model to hold comment information
// The example loads an Excel workbook, walks through each worksheet and its used cells, captures any cell comments into a simple model (sheet name, address, and comment text), serializes the collection to an indented JSON string, and writes it to a file, with basic error handling for missing files and runtime exceptions.
class CommentInfo
{
    public string SheetName { get; set; }
    public string Address { get; set; }
    public string Text { get; set; }
}

class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "comments.json";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Collect comments from all worksheets
            var comments = new List<CommentInfo>();
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all cells that contain data
                var cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxColumn = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxColumn; col++)
                    {
                        var cell = cells[row, col];
                        if (cell?.Comment != null && !string.IsNullOrEmpty(cell.Comment.Note))
                        {
                            comments.Add(new CommentInfo
                            {
                                SheetName = sheet.Name,
                                Address = cell.Name,
                                Text = cell.Comment.Note
                            });
                        }
                    }
                }
            }

            // Serialize comments to JSON with indentation
            var json = JsonSerializer.Serialize(comments, new JsonSerializerOptions { WriteIndented = true });

            // Write JSON to a file
            File.WriteAllText(outputPath, json);
            Console.WriteLine($"Comments extracted and saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
