using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

class XpsJsonToExcel
{
    static void Main()
    {
        // Paths for input JSON (representing XPS data) and output Excel file
        string jsonPath = "input.json";
        string excelPath = "output.xlsx";

        // Read the entire JSON content from file
        string jsonContent = File.ReadAllText(jsonPath);

        // Parse JSON into a list of dictionaries (each dictionary = one row)
        var rows = new List<Dictionary<string, JsonElement>>();
        try
        {
            using (JsonDocument doc = JsonDocument.Parse(jsonContent))
            {
                if (doc.RootElement.ValueKind == JsonValueKind.Array)
                {
                    foreach (JsonElement element in doc.RootElement.EnumerateArray())
                    {
                        if (element.ValueKind == JsonValueKind.Object)
                        {
                            var dict = new Dictionary<string, JsonElement>();
                            foreach (JsonProperty prop in element.EnumerateObject())
                            {
                                dict[prop.Name] = prop.Value;
                            }
                            rows.Add(dict);
                        }
                    }
                }
            }
        }
        catch (JsonException)
        {
            // Invalid JSON – treat as empty data set
            rows = new List<Dictionary<string, JsonElement>>();
        }

        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // If JSON is empty or invalid, just save an empty workbook
        if (rows.Count == 0)
        {
            workbook.Save(excelPath, SaveFormat.Xlsx);
            Console.WriteLine("JSON file contained no data or was invalid. Empty Excel file created.");
            return;
        }

        // Determine column headers from the keys of the first row
        List<string> columnNames = new List<string>(rows[0].Keys);

        // Write header row (row 0)
        for (int i = 0; i < columnNames.Count; i++)
        {
            sheet.Cells[0, i].PutValue(columnNames[i]);
        }

        // Write each data row starting from row index 1
        for (int r = 0; r < rows.Count; r++)
        {
            var row = rows[r];
            for (int c = 0; c < columnNames.Count; c++)
            {
                string colName = columnNames[c];
                if (row.TryGetValue(colName, out JsonElement value))
                {
                    switch (value.ValueKind)
                    {
                        case JsonValueKind.Number:
                            if (value.TryGetInt64(out long l))
                                sheet.Cells[r + 1, c].PutValue(l);
                            else if (value.TryGetDouble(out double d))
                                sheet.Cells[r + 1, c].PutValue(d);
                            break;
                        case JsonValueKind.String:
                            sheet.Cells[r + 1, c].PutValue(value.GetString());
                            break;
                        case JsonValueKind.True:
                        case JsonValueKind.False:
                            sheet.Cells[r + 1, c].PutValue(value.GetBoolean());
                            break;
                        default:
                            sheet.Cells[r + 1, c].PutValue(value.ToString());
                            break;
                    }
                }
            }
        }

        // Save the populated workbook as XLSX
        workbook.Save(excelPath, SaveFormat.Xlsx);
        Console.WriteLine($"Excel workbook successfully created at: {excelPath}");
    }
}