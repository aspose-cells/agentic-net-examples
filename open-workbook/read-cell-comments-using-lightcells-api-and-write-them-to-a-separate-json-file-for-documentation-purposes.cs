// Title: Read Excel cell comments with Aspose.Cells LightCells API and export to JSON (C#)
// Description: Loads an XLSX workbook using Aspose.Cells LightCells, scans every worksheet's used range, extracts each comment's author, note and cell address (A1 notation), and writes the collection to a formatted JSON file while handling missing files and creating the output folder automatically.
// Keywords: Aspose.Cells LightCells read comments | C# extract Excel comments | export Excel comments to JSON | cell notes extraction Aspose | Aspose.Cells comment API | Excel documentation JSON
// Common Searches: how to read cell comments with Aspose.Cells C# | export Excel comments as JSON using Aspose | Aspose.Cells LightCells comment extraction example | C# code to list all worksheet comments in JSON
// Developer Intent: Retrieve every comment from an Excel workbook and save the details (sheet, cell, author, note) into a JSON document.
// Use Cases: Create a searchable documentation file of all annotations in a financial model. | Migrate legacy Excel comment data into a JSON‑based issue tracker or knowledge base. | Capture a snapshot of worksheet notes before performing bulk data processing.
// AI Prompts: Generate C# code that uses Aspose.Cells LightCells to read all cell comments and output a JSON array with sheet name, cell address, author, and note. | Add robust error handling for missing input files, permission errors, and empty comment collections when exporting to JSON. | Show how to ensure the output directory exists and format the JSON with indentation for readability.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// Loads an XLSX workbook using Aspose.Cells LightCells, scans every worksheet's used range, extracts each comment's author, note and cell address (A1 notation), and writes the collection to a formatted JSON file while handling missing files and creating the output folder automatically.
class Program
{
    static void Main()
    {
        // Input Excel file containing comments
        string inputPath = "CommentsDemo.xlsx";

        // Output JSON file that will store extracted comments
        string jsonPath = "comments.json";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook (read‑only mode is not required; simply load the file)
            var loadOptions = new LoadOptions(LoadFormat.Xlsx);
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Collection to hold comment information
            var commentInfos = new List<object>();

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the used range of the worksheet
                var cells = sheet.Cells;
                int startRow = cells.MinRow;
                int startColumn = cells.MinColumn;
                int endRow = cells.MaxRow;
                int endColumn = cells.MaxColumn;

                // Iterate through each cell in the used range
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startColumn; col <= endColumn; col++)
                    {
                        // Retrieve comment for the current cell (if any)
                        Comment comment = sheet.Comments[row, col];
                        if (comment != null)
                        {
                            // Convert row/column indices to Excel cell name (e.g., "A1")
                            string cellName = CellsHelper.CellIndexToName(row, col);

                            // Store desired comment details
                            commentInfos.Add(new
                            {
                                Sheet = sheet.Name,
                                Cell = cellName,
                                Author = comment.Author,
                                Note = comment.Note
                            });
                        }
                    }
                }
            }

            // Serialize the comment collection to formatted JSON
            var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
            string json = JsonSerializer.Serialize(commentInfos, jsonOptions);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(jsonPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Write JSON to the output file
            File.WriteAllText(jsonPath, json);
            Console.WriteLine($"Comments extracted successfully to '{jsonPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
