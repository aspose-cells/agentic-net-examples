using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

class JsonToXlsmConverter
{
    static void Main()
    {
        // Example JSON payload representing a simple 2‑dimensional array.
        // In a real scenario replace this string with the actual JSON input.
        string jsonPayload = @"
        [
            [""Name"", ""Age"", ""Country""],
            [""Alice"", 30, ""USA""],
            [""Bob"", 25, ""UK""],
            [""Charlie"", 28, ""Canada""]
        ]";

        // Parse the JSON payload.
        JsonDocument doc = JsonDocument.Parse(jsonPayload);
        JsonElement root = doc.RootElement;

        // Create a new workbook (default format is Xlsx, we will save as Xlsm later).
        Workbook workbook = new Workbook();                     // uses Workbook() constructor rule
        Worksheet sheet = workbook.Worksheets[0];               // default first worksheet

        // Populate the worksheet with the data from the JSON array.
        // Expecting the JSON to be an array of rows, each row an array of cell values.
        int rowIndex = 0;
        foreach (JsonElement row in root.EnumerateArray())
        {
            int colIndex = 0;
            foreach (JsonElement cell in row.EnumerateArray())
            {
                // Determine the JSON value type and put the appropriate value into the cell.
                switch (cell.ValueKind)
                {
                    case JsonValueKind.Number:
                        if (cell.TryGetInt32(out int intVal))
                            sheet.Cells[rowIndex, colIndex].PutValue(intVal);
                        else if (cell.TryGetDouble(out double dblVal))
                            sheet.Cells[rowIndex, colIndex].PutValue(dblVal);
                        break;
                    case JsonValueKind.String:
                        sheet.Cells[rowIndex, colIndex].PutValue(cell.GetString());
                        break;
                    case JsonValueKind.True:
                    case JsonValueKind.False:
                        sheet.Cells[rowIndex, colIndex].PutValue(cell.GetBoolean());
                        break;
                    default:
                        // For null or other types, leave the cell empty.
                        break;
                }
                colIndex++;
            }
            rowIndex++;
        }

        // Save the workbook as an XLSM file (macro‑enabled format) using default settings.
        string outputPath = "output.xlsm";
        workbook.Save(outputPath, SaveFormat.Xlsm);             // uses Save(string, SaveFormat) rule

        Console.WriteLine($"Workbook created and saved to '{outputPath}'.");
    }
}