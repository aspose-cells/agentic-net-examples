using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using Aspose.Cells.Utility;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsJsonExportWithComments
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with a header row
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Alice");
            sheet.Cells["B3"].PutValue(25);

            // Add comments to some cells
            int commentIndex1 = sheet.Comments.Add(1, 0); // A2
            sheet.Comments[commentIndex1].Note = "Employee name";

            int commentIndex2 = sheet.Comments.Add(2, 1); // B3
            sheet.Comments[commentIndex2].Note = "Age in years";

            // Define the range to export (including header)
            AsposeRange exportRange = sheet.Cells.CreateRange("A1:B3");

            // Configure JSON export options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportAsString = true,
                Indent = "    "
            };

            // Export the range to a JSON string
            string rawJson = JsonUtility.ExportRangeToJson(exportRange, jsonOptions);

            // Parse the JSON into a mutable node structure
            JsonArray rows = JsonNode.Parse(rawJson)!.AsArray();

            // Retrieve header names (first row of the range)
            string[] headers = new string[exportRange.ColumnCount];
            for (int col = 0; col < exportRange.ColumnCount; col++)
            {
                Cell headerCell = sheet.Cells[exportRange.FirstRow, exportRange.FirstColumn + col];
                headers[col] = headerCell.StringValue;
            }

            // Iterate over data rows (skip header row)
            for (int rowIdx = 0; rowIdx < rows.Count; rowIdx++)
            {
                JsonObject rowObject = rows[rowIdx]!.AsObject();

                // The JSON rows correspond to the data rows (A2:B3), so map to worksheet rows
                int worksheetRow = exportRange.FirstRow + 1 + rowIdx; // +1 to skip header

                for (int col = 0; col < exportRange.ColumnCount; col++)
                {
                    // Check if a comment exists for this cell
                    Comment comment = sheet.Comments[worksheetRow, exportRange.FirstColumn + col];
                    if (comment != null && !string.IsNullOrEmpty(comment.Note))
                    {
                        // Add a new JSON property for the comment.
                        // Property name: "<Header>_Comment"
                        string commentPropertyName = headers[col] + "_Comment";
                        rowObject[commentPropertyName] = comment.Note;
                    }
                }
            }

            // Convert the modified JSON back to string
            string jsonWithComments = rows.ToJsonString(new JsonSerializerOptions { WriteIndented = true });

            // Output the final JSON
            Console.WriteLine("Exported JSON with comments:");
            Console.WriteLine(jsonWithComments);

            // (Optional) Save the JSON to a file
            File.WriteAllText("ExportedDataWithComments.json", jsonWithComments);
        }
    }
}