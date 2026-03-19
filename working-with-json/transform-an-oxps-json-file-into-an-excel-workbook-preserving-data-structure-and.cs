using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace OXpsJsonToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input OXPS JSON file path (replace with actual path)
            string jsonFilePath = "input.json";

            // Output Excel file path
            string excelFilePath = "output.xlsx";

            // Ensure the JSON file exists
            if (!File.Exists(jsonFilePath))
            {
                Console.WriteLine($"JSON file not found: {jsonFilePath}");
                return;
            }

            // Load JSON content
            string jsonContent = File.ReadAllText(jsonFilePath);

            JsonDocument doc;
            try
            {
                doc = JsonDocument.Parse(jsonContent);
            }
            catch (JsonException ex)
            {
                Console.WriteLine($"Failed to parse JSON: {ex.Message}");
                return;
            }

            using (doc)
            {
                JsonElement root = doc.RootElement;

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Determine if the root element is an array (typical tabular JSON)
                if (root.ValueKind == JsonValueKind.Array)
                {
                    int rowIndex = 0;
                    string[] headers = null;

                    foreach (JsonElement element in root.EnumerateArray())
                    {
                        if (element.ValueKind != JsonValueKind.Object)
                            continue; // Skip non-object entries

                        // On first iteration, extract headers from property names
                        if (headers == null)
                        {
                            int col = 0;
                            int headerCount = element.EnumerateObject().Count();
                            headers = new string[headerCount];
                            foreach (JsonProperty prop in element.EnumerateObject())
                            {
                                headers[col] = prop.Name;
                                // Write header to first row
                                sheet.Cells[rowIndex, col].PutValue(prop.Name);
                                col++;
                            }
                            rowIndex++; // Move to data rows
                        }

                        // Write data row
                        int dataCol = 0;
                        foreach (JsonProperty prop in element.EnumerateObject())
                        {
                            JsonElement value = prop.Value;
                            // Handle basic JSON value types
                            switch (value.ValueKind)
                            {
                                case JsonValueKind.String:
                                    sheet.Cells[rowIndex, dataCol].PutValue(value.GetString());
                                    break;
                                case JsonValueKind.Number:
                                    if (value.TryGetInt64(out long l))
                                        sheet.Cells[rowIndex, dataCol].PutValue(l);
                                    else if (value.TryGetDouble(out double d))
                                        sheet.Cells[rowIndex, dataCol].PutValue(d);
                                    break;
                                case JsonValueKind.True:
                                case JsonValueKind.False:
                                    sheet.Cells[rowIndex, dataCol].PutValue(value.GetBoolean());
                                    break;
                                case JsonValueKind.Null:
                                    sheet.Cells[rowIndex, dataCol].PutValue(string.Empty);
                                    break;
                                default:
                                    // For complex types, store the raw JSON string
                                    sheet.Cells[rowIndex, dataCol].PutValue(value.GetRawText());
                                    break;
                            }
                            dataCol++;
                        }
                        rowIndex++;
                    }
                }
                else if (root.ValueKind == JsonValueKind.Object)
                {
                    // If the JSON root is a single object, treat its properties as a single row
                    int col = 0;
                    foreach (JsonProperty prop in root.EnumerateObject())
                    {
                        // Header
                        sheet.Cells[0, col].PutValue(prop.Name);
                        // Value
                        JsonElement value = prop.Value;
                        switch (value.ValueKind)
                        {
                            case JsonValueKind.String:
                                sheet.Cells[1, col].PutValue(value.GetString());
                                break;
                            case JsonValueKind.Number:
                                if (value.TryGetInt64(out long l))
                                    sheet.Cells[1, col].PutValue(l);
                                else if (value.TryGetDouble(out double d))
                                    sheet.Cells[1, col].PutValue(d);
                                break;
                            case JsonValueKind.True:
                            case JsonValueKind.False:
                                sheet.Cells[1, col].PutValue(value.GetBoolean());
                                break;
                            case JsonValueKind.Null:
                                sheet.Cells[1, col].PutValue(string.Empty);
                                break;
                            default:
                                sheet.Cells[1, col].PutValue(value.GetRawText());
                                break;
                        }
                        col++;
                    }
                }
                else
                {
                    Console.WriteLine("Unsupported JSON structure for conversion.");
                    return;
                }

                // Save the workbook to the specified Excel file
                workbook.Save(excelFilePath, SaveFormat.Xlsx);
                Console.WriteLine($"Conversion completed. Excel file saved at: {excelFilePath}");
            }
        }
    }
}