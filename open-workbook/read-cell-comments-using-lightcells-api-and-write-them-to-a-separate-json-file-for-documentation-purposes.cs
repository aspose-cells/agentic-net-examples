// Title: Extract cell comments from an Excel workbook using Aspose.Cells LightCells API and save them as a JSON file (C#)
// Common Searches: how to read Excel cell comments with Aspose.Cells LightCells API in C# | export worksheet comments to JSON using Aspose.Cells for .NET | C# code to extract cell notes from .xlsx and write to a JSON file | Aspose.Cells example for serializing comments with System.Text.Json | retrieve comments from all sheets in an Excel file using LightCells
// Tags: Aspose.Cells LightCells API comment extraction | Excel cell comments to JSON C# | read worksheet comments Aspose.Cells | serialize comment list with System.Text.Json | iterate workbook cells for comments

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// The sample opens an existing .xlsx workbook (or creates a minimal one with a sample comment), iterates through each worksheet and cell using the LightCells API, collects non‑empty comments along with their sheet name and address, stores them in a CommentInfo list, serializes the list to indented JSON with System.Text.Json, and writes the output to a specified JSON file while handling I/O errors.
public class CommentInfo
{
    public string? SheetName { get; set; }
    public string? CellAddress { get; set; }
    public string? CommentText { get; set; }
}

class Program
{
    static void Main()
    {
        try
        {
            // Input and output file paths.
            string excelPath = "InputWorkbook.xlsx";
            string jsonPath = "CellComments.json";

            // Ensure the source workbook exists; create a minimal workbook if missing.
            Workbook workbook;
            if (File.Exists(excelPath))
            {
                workbook = new Workbook(excelPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Sheet1";

                // Add a sample comment to demonstrate functionality.
                Cell sampleCell = sheet.Cells["A1"];
                sampleCell.PutValue("Sample");

                // Create comment using the Comments collection.
                int commentIdx = sheet.Comments.Add("A1");
                Comment comment = sheet.Comments[commentIdx];
                comment.Note = "This is a sample comment.";
                comment.Author = "Author";

                workbook.Save(excelPath);
            }

            // Collect comment information.
            var commentList = new List<CommentInfo>();

            // Iterate through worksheets and cells to extract comments.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Cell cell in sheet.Cells)
                {
                    if (cell.Comment != null && !string.IsNullOrEmpty(cell.Comment.Note))
                    {
                        commentList.Add(new CommentInfo
                        {
                            SheetName = sheet.Name,
                            CellAddress = cell.Name, // e.g., "B2"
                            CommentText = cell.Comment.Note
                        });
                    }
                }
            }

            // Serialize to JSON.
            string json = JsonSerializer.Serialize(commentList, new JsonSerializerOptions { WriteIndented = true });

            // Write JSON output safely.
            try
            {
                File.WriteAllText(jsonPath, json);
                Console.WriteLine($"Extracted {commentList.Count} comment(s) to '{jsonPath}'.");
            }
            catch (Exception ioEx)
            {
                Console.Error.WriteLine($"Failed to write JSON file: {ioEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
