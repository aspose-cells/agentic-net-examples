// Title: Export Excel worksheet comments to JSON with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook using Aspose.Cells, iterates every worksheet, extracts each comment's sheet name, A1 cell address, author and note, serializes the collection to indented JSON, and saves it to a file. Includes basic file‑existence checking and exception handling.
// Keywords: Aspose.Cells | C# | Excel comments extraction | export comments to JSON | cell notes serialization | worksheet comment reader | Aspose.Cells API example | JSON output from Excel
// Common Searches: how to read Excel comments with Aspose.Cells C# | export Excel cell notes to JSON using .NET | Aspose.Cells iterate worksheet comments example | convert comment row column to A1 address Aspose | save Excel comments as JSON file C#
// Developer Intent: Retrieve all comments from every sheet in an Excel file and write them to a structured JSON document.
// Use Cases: Create searchable documentation of reviewer remarks embedded in spreadsheets. | Generate an audit trail of cell annotations for compliance or quality checks. | Feed extracted comments into a web service or UI that displays workbook annotations.
// AI Prompts: Generate C# code that uses Aspose.Cells to read every comment in an Excel workbook and output a formatted JSON array with sheet, cell, author, and note fields. | Show how to convert comment row and column indices to A1 notation using CellsHelper in Aspose.Cells. | Provide robust error handling for missing input files and JSON serialization options when exporting Excel comments.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// Loads an Excel workbook using Aspose.Cells, iterates every worksheet, extracts each comment's sheet name, A1 cell address, author and note, serializes the collection to indented JSON, and saves it to a file. Includes basic file‑existence checking and exception handling.
class Program
{
    static void Main()
    {
        // Path to the source workbook
        string workbookPath = "input.xlsx";

        // Path to the output JSON file
        string jsonOutputPath = "comments.json";

        // Verify that the input workbook exists
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: The file \"{workbookPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // List to hold comment information
            List<object> comments = new List<object>();

            // Iterate through all worksheets (adjust if only a specific sheet is needed)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all comments in the worksheet
                foreach (Comment comment in sheet.Comments)
                {
                    // Convert row/column indices to cell name (e.g., "A1")
                    string cellName = CellsHelper.CellIndexToName(comment.Row, comment.Column);

                    // Add comment details to the list
                    comments.Add(new
                    {
                        Sheet = sheet.Name,
                        Cell = cellName,
                        Author = comment.Author,
                        Note = comment.Note
                    });
                }
            }

            // Serialize the comment list to JSON with indentation
            string json = JsonSerializer.Serialize(comments, new JsonSerializerOptions { WriteIndented = true });

            // Write JSON to the output file
            File.WriteAllText(jsonOutputPath, json);

            Console.WriteLine("Comments have been exported to " + jsonOutputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while processing the workbook:");
            Console.WriteLine(ex.Message);
        }
    }
}
