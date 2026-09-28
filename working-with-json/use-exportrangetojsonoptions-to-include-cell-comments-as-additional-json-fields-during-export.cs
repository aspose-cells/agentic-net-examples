// Title: Export an Excel range to JSON and include each cell's comment as a separate property using Aspose.Cells ExportRangeToJsonOptions (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells ExportRangeToJsonOptions to export a worksheet range to JSON, adding a *_Comment field for cells that contain comments. | Show how to configure ExportRangeToJsonOptions so that cell comment notes are emitted as additional JSON properties while preserving original column headers. | Create an example that reads an Excel sheet, extracts data and comments, and produces an indented JSON array where each data field has a corresponding comment field when a comment exists.
// Common Searches: how to export excel data to json with comments using aspose.cells c# | asp.net export worksheet range to json including cell notes | c# aspose.cells ExportRangeToJsonOptions add comment fields to json output | include excel cell comments as separate json properties with Aspose.Cells
// Tags: exportrangejsonoptions include comments c# | aspose.cells export excel to json with annotations | excel cell comment to json field asp.net | c# convert worksheet range to json with comment metadata | aspose.cells json export range with notes

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// // Demonstrates creating a workbook, adding data and comments, then building a JSON array where each column value is paired with an extra *_Comment property that contains the cell's comment.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate header and data rows
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Jane");
            sheet.Cells["B3"].PutValue(25);

            // Add comments to cells (row, column are zero‑based). Provide an author for the comment.
            sheet.Comments[1, 0].Author = "Author";
            sheet.Comments[1, 0].Note = "First name";          // A2
            sheet.Comments[1, 1].Author = "Author";
            sheet.Comments[1, 1].Note = "Age in years";        // B2
            sheet.Comments[2, 0].Author = "Author";
            sheet.Comments[2, 0].Note = "First name";          // A3
            sheet.Comments[2, 1].Author = "Author";
            sheet.Comments[2, 1].Note = "Age in years";        // B3

            // Build JSON manually, including comments as extra fields
            var headers = new List<string>();
            for (int c = 0; c < 2; c++)
                headers.Add(sheet.Cells[0, c].StringValue);

            var records = new List<Dictionary<string, object>>();
            for (int r = 1; r <= 2; r++) // data rows (A2:B3)
            {
                var rowDict = new Dictionary<string, object>();
                for (int c = 0; c < 2; c++)
                {
                    var value = sheet.Cells[r, c].Value;
                    rowDict[headers[c]] = value;

                    // Append comment if present
                    var comment = sheet.Comments[r, c];
                    if (comment != null && !string.IsNullOrEmpty(comment.Note))
                        rowDict[$"{headers[c]}_Comment"] = comment.Note;
                }
                records.Add(rowDict);
            }

            string json = JsonSerializer.Serialize(records, new JsonSerializerOptions { WriteIndented = true });

            // Save JSON to file
            string outputPath = "ExportedData.json";
            File.WriteAllText(outputPath, json);
            Console.WriteLine($"JSON data exported successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
