using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using Aspose.Cells.Utility;

class ExportRangeToJsonWithComments
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (including a header row)
            cells["A1"].PutValue("Name");
            cells["B1"].PutValue("Age");
            cells["A2"].PutValue("John");
            cells["B2"].PutValue(30);
            cells["A3"].PutValue("Alice");
            cells["B3"].PutValue(25);

            // Add comments to some cells
            int commentIdx1 = sheet.Comments.Add("A2");
            sheet.Comments[commentIdx1].Note = "Employee name";
            int commentIdx2 = sheet.Comments.Add("B3");
            sheet.Comments[commentIdx2].Note = "Age in years";

            // Define the range to export (including header)
            Aspose.Cells.Range exportRange = cells.CreateRange("A1:B3");

            // Configure ExportRangeToJsonOptions
            ExportRangeToJsonOptions jsonOptions = new ExportRangeToJsonOptions
            {
                HasHeaderRow = true,
                ExportEmptyCells = true,
                ExportAsString = false,
                Indent = "  " // pretty‑print with two spaces
            };

            // Export the range to a JSON string
            string json = JsonUtility.ExportRangeToJson(exportRange, jsonOptions);

            // Build a dictionary of cell address -> comment text for quick lookup
            Dictionary<string, string> commentMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (Comment comment in sheet.Comments)
            {
                // comment.CommentShape?.Name contains the cell name (e.g., "A2")
                string cellName = comment.CommentShape?.Name;
                if (!string.IsNullOrEmpty(cellName) && !string.IsNullOrEmpty(comment.Note))
                {
                    commentMap[cellName] = comment.Note;
                }
            }

            // Parse the exported JSON so we can inject comment fields
            JsonNode rootNode = JsonNode.Parse(json);
            if (rootNode is not JsonArray rowsArray)
            {
                Console.WriteLine("Unexpected JSON format.");
                return;
            }

            // Retrieve header names from the first row of the range
            List<string> headers = new List<string>();
            for (int col = exportRange.FirstColumn; col <= exportRange.FirstColumn + exportRange.ColumnCount - 1; col++)
            {
                object headerVal = cells[exportRange.FirstRow, col].Value;
                headers.Add(headerVal?.ToString() ?? $"Column{col}");
            }

            // Iterate over data rows and add comment fields where applicable
            for (int i = 0; i < rowsArray.Count; i++)
            {
                if (rowsArray[i] is not JsonObject rowObj) continue;

                int excelRow = exportRange.FirstRow + 1 + i; // data rows start after header
                for (int colOffset = 0; colOffset < headers.Count; colOffset++)
                {
                    int excelCol = exportRange.FirstColumn + colOffset;
                    string cellAddress = CellsHelper.CellIndexToName(excelRow, excelCol); // e.g., "A2"

                    if (commentMap.TryGetValue(cellAddress, out string commentText))
                    {
                        // Add a new field named "<Header>_Comment" with the comment text
                        string commentFieldName = $"{headers[colOffset]}_Comment";
                        rowObj[commentFieldName] = commentText;
                    }
                }
            }

            // Serialize the enriched JSON back to a string with indentation
            string enrichedJson = rootNode.ToJsonString(new JsonSerializerOptions { WriteIndented = true });

            // Output the final JSON
            Console.WriteLine("Exported JSON with comments:");
            Console.WriteLine(enrichedJson);

            // Optionally, save the JSON to a file
            File.WriteAllText("ExportedWithComments.json", enrichedJson);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}