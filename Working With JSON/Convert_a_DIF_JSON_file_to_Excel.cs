using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

class DifJsonToExcel
{
    static void Main()
    {
        string jsonFilePath = "data.json";
        string excelFilePath = "result.xlsx";

        string jsonContent = File.ReadAllText(jsonFilePath);
        JsonElement root = JsonSerializer.Deserialize<JsonElement>(jsonContent);

        JsonElement rowsElement;
        if (root.ValueKind == JsonValueKind.Array)
        {
            rowsElement = root;
        }
        else if (root.ValueKind == JsonValueKind.Object &&
                 root.TryGetProperty("data", out JsonElement dataProp) &&
                 dataProp.ValueKind == JsonValueKind.Array)
        {
            rowsElement = dataProp;
        }
        else if (root.ValueKind == JsonValueKind.Object)
        {
            // Treat the whole object as a single row
            rowsElement = root;
        }
        else
        {
            Console.WriteLine("Invalid JSON format. Expected an array of rows or an object with a 'data' array.");
            return;
        }

        // Prepare a list of row elements to iterate uniformly
        List<JsonElement> rows = new List<JsonElement>();
        if (rowsElement.ValueKind == JsonValueKind.Array)
        {
            foreach (JsonElement el in rowsElement.EnumerateArray())
                rows.Add(el);
        }
        else // Object treated as a single row
        {
            rows.Add(rowsElement);
        }

        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        int rowIndex = 0;
        foreach (JsonElement row in rows)
        {
            int colIndex = 0;
            if (row.ValueKind == JsonValueKind.Array)
            {
                foreach (JsonElement cell in row.EnumerateArray())
                {
                    WriteCell(sheet, rowIndex, colIndex, cell);
                    colIndex++;
                }
            }
            else if (row.ValueKind == JsonValueKind.Object)
            {
                foreach (JsonProperty prop in row.EnumerateObject())
                {
                    WriteCell(sheet, rowIndex, colIndex, prop.Value);
                    colIndex++;
                }
            }
            rowIndex++;
        }

        workbook.Save(excelFilePath, SaveFormat.Xlsx);
        Console.WriteLine($"Conversion completed: '{jsonFilePath}' → '{excelFilePath}'");
    }

    static void WriteCell(Worksheet sheet, int row, int col, JsonElement cell)
    {
        switch (cell.ValueKind)
        {
            case JsonValueKind.Number:
                if (cell.TryGetInt64(out long l))
                    sheet.Cells[row, col].PutValue(l);
                else if (cell.TryGetDouble(out double d))
                    sheet.Cells[row, col].PutValue(d);
                break;
            case JsonValueKind.String:
                sheet.Cells[row, col].PutValue(cell.GetString());
                break;
            case JsonValueKind.True:
            case JsonValueKind.False:
                sheet.Cells[row, col].PutValue(cell.GetBoolean());
                break;
            default:
                // Leave cell empty for null or unsupported types
                break;
        }
    }
}