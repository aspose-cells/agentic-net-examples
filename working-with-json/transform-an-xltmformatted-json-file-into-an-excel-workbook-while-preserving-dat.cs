using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the JSON file
        string jsonPath = "input.json";

        // Read the JSON content
        string jsonContent = File.ReadAllText(jsonPath);

        // Parse JSON and convert to list of dictionaries
        List<Dictionary<string, JsonElement>> rows = new List<Dictionary<string, JsonElement>>();
        try
        {
            var deserialized = JsonSerializer.Deserialize<List<Dictionary<string, JsonElement>>>(jsonContent);
            if (deserialized != null)
                rows = deserialized;
        }
        catch
        {
            Console.WriteLine("Failed to parse JSON. Ensure the file contains a valid JSON array of objects.");
        }

        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Data";

        // If there are no rows, exit early
        if (rows == null || rows.Count == 0)
        {
            Console.WriteLine("No data found in JSON.");
            return;
        }

        // Determine column headers from the first row's keys
        var headers = new List<string>(rows[0].Keys);

        // Write headers to the first row
        for (int col = 0; col < headers.Count; col++)
        {
            sheet.Cells[0, col].PutValue(headers[col]);
        }

        // Write data rows starting from row index 1
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            var rowDict = rows[rowIndex];
            for (int col = 0; col < headers.Count; col++)
            {
                string header = headers[col];
                if (rowDict.TryGetValue(header, out JsonElement element))
                {
                    switch (element.ValueKind)
                    {
                        case JsonValueKind.Number:
                            if (element.TryGetInt64(out long l))
                                sheet.Cells[rowIndex + 1, col].PutValue((double)l);
                            else if (element.TryGetDouble(out double d))
                                sheet.Cells[rowIndex + 1, col].PutValue(d);
                            else
                                sheet.Cells[rowIndex + 1, col].PutValue(element.GetRawText());
                            break;
                        case JsonValueKind.String:
                            string str = element.GetString();
                            if (DateTime.TryParse(str, out DateTime dt))
                                sheet.Cells[rowIndex + 1, col].PutValue(dt);
                            else
                                sheet.Cells[rowIndex + 1, col].PutValue(str);
                            break;
                        case JsonValueKind.True:
                        case JsonValueKind.False:
                            sheet.Cells[rowIndex + 1, col].PutValue(element.GetBoolean());
                            break;
                        case JsonValueKind.Null:
                            sheet.Cells[rowIndex + 1, col].PutValue(string.Empty);
                            break;
                        default:
                            sheet.Cells[rowIndex + 1, col].PutValue(element.ToString());
                            break;
                    }
                }
                else
                {
                    // Empty cell if key missing
                    sheet.Cells[rowIndex + 1, col].PutValue(string.Empty);
                }
            }
        }

        // Auto‑fit columns for better readability
        sheet.AutoFitColumns();

        // Save the workbook to an Excel file
        string outputPath = "output.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);

        Console.WriteLine($"Conversion completed. Excel file saved to '{outputPath}'.");
    }
}