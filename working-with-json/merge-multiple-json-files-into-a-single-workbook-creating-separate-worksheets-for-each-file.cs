// Title: Combine several JSON files into one Excel workbook with a separate worksheet for each file using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells how to add a worksheet for each JSON file and merge them into one workbook | import multiple JSON arrays into separate sheets of an Excel file using Aspose.Cells .NET | merge JSON data files into a single Excel workbook programmatically with Aspose.Cells | create workbook with sheet name derived from JSON filename using Aspose.Cells C# | auto‑fit columns after importing JSON data with Aspose.Cells in .NET
// Tags: Aspose.Cells import JSON array to worksheet | C# generate Excel workbook from multiple JSON files | Aspose.Cells create separate sheets per JSON source | Aspose.Cells auto‑fit columns programmatically | System.Text.Json parse JSON for Excel export | Aspose.Cells workbook.Save Xlsx format

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

// The example reads a list of JSON files, creates a new workbook, clears the default sheet, and for each existing JSON file adds a worksheet named after the file (without extension). It parses the JSON array with System.Text.Json, writes column headers from the first object, fills subsequent rows with values, auto‑fits columns, and finally saves the combined workbook as MergedWorkbook.xlsx while handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        // Paths of JSON files to be merged
        List<string> jsonFiles = new List<string>
        {
            "data1.json",
            "data2.json",
            "data3.json"
        };

        // Create a new workbook and clear default worksheets
        Workbook workbook = new Workbook();
        workbook.Worksheets.Clear();

        foreach (string jsonPath in jsonFiles)
        {
            try
            {
                // Ensure the JSON file exists
                if (!File.Exists(jsonPath))
                {
                    Console.WriteLine($"File not found: {jsonPath}");
                    continue;
                }

                // Read JSON content
                string jsonContent = File.ReadAllText(jsonPath);

                // Use the file name (without extension) as the worksheet name
                string sheetName = Path.GetFileNameWithoutExtension(jsonPath);
                Worksheet sheet = workbook.Worksheets.Add(sheetName);

                // Import JSON data into the worksheet
                ImportJsonToWorksheet(jsonContent, sheet);

                // Auto‑fit columns for readability
                sheet.AutoFitColumns();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{jsonPath}': {ex.Message}");
            }
        }

        // Save the combined workbook
        try
        {
            workbook.Save("MergedWorkbook.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved as 'MergedWorkbook.xlsx'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }

    private static void ImportJsonToWorksheet(string jsonContent, Worksheet sheet)
    {
        // Parse JSON
        using JsonDocument doc = JsonDocument.Parse(jsonContent);
        if (doc.RootElement.ValueKind != JsonValueKind.Array)
            throw new InvalidOperationException("Root JSON element must be an array.");

        JsonElement.ArrayEnumerator rows = doc.RootElement.EnumerateArray();

        // Determine column headers from the first object
        List<string> headers = new List<string>();
        bool headersDefined = false;
        int currentRow = 0;

        foreach (JsonElement row in rows)
        {
            if (!headersDefined)
            {
                foreach (JsonProperty prop in row.EnumerateObject())
                    headers.Add(prop.Name);
                // Write headers
                for (int col = 0; col < headers.Count; col++)
                    sheet.Cells[currentRow, col].PutValue(headers[col]);
                currentRow++;
                headersDefined = true;
            }

            // Write row values
            int colIndex = 0;
            foreach (string header in headers)
            {
                if (row.TryGetProperty(header, out JsonElement value))
                {
                    // Store the raw value; Aspose.Cells will handle conversion
                    sheet.Cells[currentRow, colIndex].PutValue(GetJsonElementValue(value));
                }
                else
                {
                    sheet.Cells[currentRow, colIndex].PutValue(string.Empty);
                }
                colIndex++;
            }
            currentRow++;
        }
    }

    /// <summary>
    /// Converts a JsonElement to a .NET primitive suitable for Aspose.Cells.
    /// </summary>
    private static object GetJsonElementValue(JsonElement element)
    {
        return element.ValueKind switch
        {
            JsonValueKind.String => element.GetString(),
            JsonValueKind.Number => element.TryGetInt64(out long l) ? (object)l :
                                    element.TryGetDouble(out double d) ? d : element.GetDecimal(),
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.Null => null,
            _ => element.GetRawText()
        };
    }
}
