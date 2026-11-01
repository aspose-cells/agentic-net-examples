// Title: Read cell comments from an Excel workbook with Aspose.Cells (C#) and export them to a formatted JSON file
// Common Searches: Aspose.Cells C# read worksheet comments and save as JSON | How to extract Excel cell notes using LightCells API in .NET | Export Excel comments to a JSON file with Aspose.Cells library | C# iterate through .xlsx comments and generate indented JSON | Serialize Aspose.Cells comment collection to JSON in a .NET application
// Tags: Aspose.Cells comment extraction | Excel comment data to JSON | C# LightCells API comment read | serialize worksheet notes | JSON generation from cell comments | comment metadata extraction .NET

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// The example loads an Excel workbook using Aspose.Cells, walks through each worksheet's Comments collection via the LightCells API, captures the sheet name, zero‑based row and column indices, and the comment text. It then serializes the gathered information into an indented JSON array and writes it to a specified file, with error handling for missing workbooks.
class Program
{
    static void Main()
    {
        // Paths for input workbook and output JSON
        string workbookPath = "input.xlsx";
        string jsonPath = "comments.json";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(workbookPath))
                throw new FileNotFoundException($"The workbook file '{workbookPath}' was not found.");

            // Load the workbook
            var workbook = new Workbook(workbookPath);

            // Collect comment information
            var commentInfos = new List<object>();

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through comments in the current worksheet
                foreach (Comment comment in sheet.Comments)
                {
                    var info = new
                    {
                        Sheet = sheet.Name,
                        Row = comment.Row,       // Zero‑based row index
                        Column = comment.Column, // Zero‑based column index
                        Text = comment.Note
                    };
                    commentInfos.Add(info);
                }
            }

            // Serialize comments to formatted JSON
            var json = JsonSerializer.Serialize(commentInfos, new JsonSerializerOptions
            {
                WriteIndented = true
            });

            // Write JSON to the output file
            File.WriteAllText(jsonPath, json);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
