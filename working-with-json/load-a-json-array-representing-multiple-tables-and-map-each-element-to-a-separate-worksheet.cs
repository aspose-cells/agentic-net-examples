// Title: Read a JSON file containing an array of tables and generate a separate Excel worksheet for each table using Aspose.Cells in C#
// AI Prompts: Generate C# code that imports a JSON file, iterates over each array element, and adds a new worksheet to an Aspose.Cells workbook for every table. | Show how to extract field names from the first object of a JSON table, write them as column headers, and fill subsequent rows in separate worksheets with Aspose.Cells. | Demonstrate saving the populated workbook as an .xlsx file and implementing error handling for missing JSON files when processing multiple tables.
// Common Searches: C# Aspose.Cells create separate worksheets from each element of a JSON array of tables | how to map JSON object fields to Excel column headers using Aspose.Cells .NET | read JSON file with multiple tables and export each table to its own sheet in Aspose.Cells | Aspose.Cells import JSON and generate multi‑sheet workbook in C# | error handling for missing JSON file when creating Excel sheets with Aspose.Cells
// Tags: import JSON array into multiple worksheets Aspose.Cells | convert JSON tables to Excel sheets using Aspose.Cells | write JSON field names as Excel column headers C# | save Aspose.Cells workbook as .xlsx from JSON data | manage absent JSON file in Aspose.Cells process

using System;
using System.IO;
using System.Collections.Generic;
using System.Text.Json;
using Aspose.Cells;

// The example reads a JSON file that holds an array of tables, verifies each element, creates a new worksheet for every table in an Aspose.Cells workbook, writes column headers derived from the first object's property names, populates the rows, handles missing file and runtime errors, and finally saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the JSON file containing an array of tables
            string jsonPath = "tables.json";

            // Verify that the JSON file exists to avoid FileNotFoundException
            if (!File.Exists(jsonPath))
            {
                Console.WriteLine($"Error: JSON file not found at path '{jsonPath}'.");
                return;
            }

            // Load the entire JSON content from the file
            string jsonContent = File.ReadAllText(jsonPath);

            // Parse the JSON as an array; each element represents a separate table
            JsonDocument doc = JsonDocument.Parse(jsonContent);
            JsonElement root = doc.RootElement;

            if (root.ValueKind != JsonValueKind.Array)
            {
                Console.WriteLine("Error: Expected a JSON array at the root.");
                return;
            }

            // Create a new workbook instance
            Workbook workbook = new Workbook();

            // Remove the default worksheet that Aspose.Cells creates
            workbook.Worksheets.Clear();

            int sheetIndex = 0;

            // Iterate over each table in the JSON array
            foreach (JsonElement tableElement in root.EnumerateArray())
            {
                if (tableElement.ValueKind != JsonValueKind.Array)
                {
                    // Skip non‑array entries
                    continue;
                }

                // Add a new worksheet for the current table
                Worksheet sheet = workbook.Worksheets[workbook.Worksheets.Add()];
                sheet.Name = $"Sheet{sheetIndex + 1}";

                JsonElement.ArrayEnumerator rowsEnum = tableElement.EnumerateArray();

                // Convert rows to a list for easier indexing
                List<JsonElement> rows = new List<JsonElement>();
                foreach (var row in rowsEnum)
                {
                    rows.Add(row);
                }

                // If the table has no rows, skip to the next one
                if (rows.Count == 0)
                {
                    sheetIndex++;
                    continue;
                }

                // Determine column headers from the first row's property names
                JsonElement firstRow = rows[0];
                List<string> columns = new List<string>();
                int colIndex = 0;
                foreach (JsonProperty prop in firstRow.EnumerateObject())
                {
                    columns.Add(prop.Name);
                    // Write header to the first row of the worksheet
                    sheet.Cells[0, colIndex].PutValue(prop.Name);
                    colIndex++;
                }

                // Write each data row into the worksheet
                for (int i = 0; i < rows.Count; i++)
                {
                    JsonElement rowObj = rows[i];
                    for (int j = 0; j < columns.Count; j++)
                    {
                        if (rowObj.TryGetProperty(columns[j], out JsonElement valueElement))
                        {
                            // Write the cell value as string, starting from row index 1 (below headers)
                            sheet.Cells[i + 1, j].PutValue(valueElement.ToString());
                        }
                    }
                }

                sheetIndex++;
            }

            // Save the workbook to an Excel file
            workbook.Save("output.xlsx");
            Console.WriteLine("Workbook saved as 'output.xlsx'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
