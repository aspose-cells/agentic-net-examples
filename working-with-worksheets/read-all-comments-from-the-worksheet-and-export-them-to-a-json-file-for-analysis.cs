// Title: Export worksheet comments to a JSON file with Aspose.Cells for .NET
// Description: Loads a workbook, reads every comment on a selected worksheet, captures the cell address, note text and author, serializes the data into indented JSON, and writes it to comments.json.
// Keywords: Aspose.Cells export comments JSON | C# read Excel comments | serialize worksheet notes .NET | extract cell comments Aspose | save Excel remarks as JSON | Aspose.Cells comment author | JSON output from Excel comments
// Common Searches: How to export Excel comments to JSON using Aspose.Cells | Aspose.Cells C# read cell notes and author | Save worksheet comments as JSON file | Export Excel cell comments programmatically | Aspose.Cells comment extraction example
// Developer Intent: Generate a JSON file containing all comments from a worksheet for downstream analysis.
// Use Cases: Create an audit log of every comment in a spreadsheet for compliance tracking. | Feed comment data into a data‑pipeline that aggregates user feedback stored in Excel. | Produce documentation that lists each cell’s note and author for project hand‑off.
// AI Prompts: Write C# code using Aspose.Cells that exports worksheet comments to a JSON file with cell address, note, and author. | Extend the export script to include each comment’s font style and color in the JSON output. | Provide a version that processes all worksheets in a workbook and groups comments by sheet name in the resulting JSON.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// Loads a workbook, reads every comment on a selected worksheet, captures the cell address, note text and author, serializes the data into indented JSON, and writes it to comments.json.
class ExportCommentsToJson
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (or any specific worksheet)
        Worksheet worksheet = workbook.Worksheets[0];

        // Prepare a list to hold comment information
        var commentsInfo = new List<object>();

        // Iterate through all comments in the worksheet
        for (int i = 0; i < worksheet.Comments.Count; i++)
        {
            // Retrieve the comment object
            Comment comment = worksheet.Comments[i];

            // Get the cell address of the comment (e.g., "A1")
            string cellAddress = CellsHelper.CellIndexToName(comment.Row, comment.Column);

            // Add comment details to the list
            commentsInfo.Add(new
            {
                Cell = cellAddress,
                Note = comment.Note,
                Author = comment.Author
            });
        }

        // Serialize the list of comments to a formatted JSON string
        var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
        string jsonOutput = JsonSerializer.Serialize(commentsInfo, jsonOptions);

        // Write the JSON string to a file
        File.WriteAllText("comments.json", jsonOutput);

        // Optional: inform the user
        Console.WriteLine("Comments have been exported to comments.json");
    }
}
