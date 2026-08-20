// Title: Export Excel Range to JSON with Cell Comments Using Aspose.Cells for .NET
// Description: Demonstrates how to export a worksheet range to JSON with Aspose.Cells' ExportRangeToJsonOptions, then enrich the output by adding a "<Header>_Comment" property for each cell that contains a comment. The example creates a workbook, inserts headers, data, and comments, exports the range, merges comment data, and prints pretty‑printed JSON.
// Keywords: Aspose.Cells ExportRangeToJsonOptions | C# export Excel to JSON | include cell comments in JSON | .NET Excel to JSON with comments | Aspose.Cells add comment fields | JSON export with annotations | pretty printed JSON from Excel
// Common Searches: Aspose.Cells export range to JSON with comments | C# add Excel cell comments to JSON output | Export Excel data and comments as JSON .NET | How to include cell notes in JSON using Aspose.Cells | Export worksheet to JSON with extra comment columns
// Developer Intent: The developer needs to convert a selected Excel range into JSON while preserving any cell comments as separate fields in the resulting objects.
// Use Cases: Create API payloads that combine data values and reviewer notes from an Excel template. | Generate audit logs that capture both cell content and associated comments for compliance reporting. | Produce configuration files where comments act as metadata alongside the actual values.
// AI Prompts: Show a reusable C# method that takes a worksheet and returns JSON with data and "<Header>_Comment" properties for all commented cells. | Explain how to handle rows that lack comments so the JSON structure stays uniform. | Provide guidance on customizing ExportRangeToJsonOptions to control indentation and string formatting.

using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Alias to avoid ambiguity with System.Range
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsJsonExportWithComments
{
    // Demonstrates how to export a worksheet range to JSON with Aspose.Cells' ExportRangeToJsonOptions, then enrich the output by adding a "<Header>_Comment" property for each cell that contains a comment. The example creates a workbook, inserts headers, data, and comments, exports the range, merges comment data, and prints pretty‑printed JSON.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Add header row
                cells["A1"].PutValue("Name");
                cells["B1"].PutValue("Age");

                // Add data rows
                cells["A2"].PutValue("John");
                cells["B2"].PutValue(30);
                cells["A3"].PutValue("Alice");
                cells["B3"].PutValue(25);

                // Add comments to some data cells and store them in a dictionary for later use
                var commentMap = new Dictionary<string, string>();

                int commentIdx = sheet.Comments.Add("A2");
                Comment commentA2 = sheet.Comments[commentIdx];
                commentA2.Note = "Employee of the month";
                commentMap["A2"] = commentA2.Note;

                commentIdx = sheet.Comments.Add("B2");
                Comment commentB2 = sheet.Comments[commentIdx];
                commentB2.Note = "Salary in USD";
                commentMap["B2"] = commentB2.Note;

                // Define the range to export (including header)
                AsposeRange exportRange = cells.CreateRange("A1:B3");

                // Configure ExportRangeToJsonOptions
                ExportRangeToJsonOptions jsonOptions = new ExportRangeToJsonOptions
                {
                    HasHeaderRow = true,
                    ExportAsString = true,
                    Indent = "  " // pretty‑print JSON
                };

                // Export the range to a JSON string
                string json = JsonUtility.ExportRangeToJson(exportRange, jsonOptions);

                // Parse the JSON into a mutable node structure
                JsonNode? rootNode = JsonNode.Parse(json);
                if (rootNode is not JsonArray jsonArray)
                {
                    Console.WriteLine("Failed to parse JSON array.");
                    return;
                }

                // Header names (same order as in the worksheet)
                List<string> headers = new List<string> { "Name", "Age" };

                // Enrich each JSON object with comment fields where applicable
                for (int i = 0; i < jsonArray.Count; i++)
                {
                    // Row index in the worksheet (header row is row 0, data starts at row 1)
                    int worksheetRow = i + 1; // because HasHeaderRow = true

                    if (jsonArray[i] is not JsonObject rowObject)
                        continue;

                    for (int col = 0; col < headers.Count; col++)
                    {
                        string cellAddress = cells[worksheetRow, col].Name; // e.g., "A2"
                        if (commentMap.TryGetValue(cellAddress, out string commentText))
                        {
                            // Add a new property named "<Header>_Comment" with the comment text
                            string commentPropertyName = $"{headers[col]}_Comment";
                            rowObject[commentPropertyName] = commentText;
                        }
                    }
                }

                // Serialize the enriched JSON back to a string
                string enrichedJson = jsonArray.ToJsonString(new JsonSerializerOptions { WriteIndented = true });

                // Output the final JSON
                Console.WriteLine("Exported JSON with comments:");
                Console.WriteLine(enrichedJson);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
