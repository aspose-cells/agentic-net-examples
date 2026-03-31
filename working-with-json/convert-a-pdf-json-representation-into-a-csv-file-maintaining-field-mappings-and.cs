using System;
using System.IO;
using System.Text.Json;
using System.Collections.Generic;
using Aspose.Cells;

class PdfJsonToCsvConverter
{
    static void Main(string[] args)
    {
        // Paths for input JSON and output CSV
        string jsonPath = "input.json";
        string csvPath = "output.csv";

        // Read JSON content from file
        string jsonContent = File.ReadAllText(jsonPath);
        JsonDocumentOptions docOptions = new JsonDocumentOptions { AllowTrailingCommas = true };
        JsonDocument doc;
        try
        {
            doc = JsonDocument.Parse(jsonContent, docOptions);
        }
        catch (JsonException ex)
        {
            Console.WriteLine($"Failed to parse JSON: {ex.Message}");
            return;
        }

        JsonElement root = doc.RootElement;

        // Ensure the JSON root is an array; if it's an object, wrap it into an array
        List<JsonElement> elements = new List<JsonElement>();
        if (root.ValueKind == JsonValueKind.Array)
        {
            foreach (JsonElement el in root.EnumerateArray())
                elements.Add(el);
        }
        else if (root.ValueKind == JsonValueKind.Object)
        {
            elements.Add(root);
        }
        else
        {
            Console.WriteLine("The JSON file does not contain an array or object at the root.");
            return;
        }

        // Parse each object into a dictionary and collect all unique field names
        List<Dictionary<string, string>> rows = new List<Dictionary<string, string>>();
        HashSet<string> headerSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (JsonElement element in elements)
        {
            if (element.ValueKind != JsonValueKind.Object) continue;

            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (JsonProperty prop in element.EnumerateObject())
            {
                string value = prop.Value.ValueKind switch
                {
                    JsonValueKind.String => prop.Value.GetString(),
                    JsonValueKind.Number => prop.Value.GetRawText(),
                    JsonValueKind.True => "true",
                    JsonValueKind.False => "false",
                    JsonValueKind.Null => string.Empty,
                    _ => prop.Value.GetRawText()
                };
                dict[prop.Name] = value;
                headerSet.Add(prop.Name);
            }
            rows.Add(dict);
        }

        // Create a deterministic order for columns (alphabetical)
        List<string> headers = new List<string>(headerSet);
        headers.Sort(StringComparer.OrdinalIgnoreCase);

        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Write header row
        for (int col = 0; col < headers.Count; col++)
            cells[0, col].PutValue(headers[col]);

        // Write data rows
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            var rowDict = rows[rowIndex];
            for (int col = 0; col < headers.Count; col++)
            {
                string header = headers[col];
                if (rowDict.TryGetValue(header, out string cellValue))
                    cells[rowIndex + 1, col].PutValue(cellValue);
                else
                    cells[rowIndex + 1, col].PutValue(string.Empty);
            }
        }

        // Save the workbook as CSV
        workbook.Save(csvPath, SaveFormat.Csv);
        Console.WriteLine($"CSV file has been created at: {csvPath}");
    }
}