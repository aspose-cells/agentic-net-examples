using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace PdfJsonToExcel
{
    class Program
    {
        static void Main()
        {
            // Path to the JSON payload extracted from a PDF
            string jsonPath = "input.json";

            // Read the entire JSON content
            string jsonContent = File.ReadAllText(jsonPath);

            // Create a new workbook
            Workbook workbook = new Workbook();

            // We'll work with the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            bool parsed = false;

            try
            {
                using JsonDocument doc = JsonDocument.Parse(jsonContent);
                JsonElement root = doc.RootElement;

                // Determine if the root element is an array (common for tabular data)
                if (root.ValueKind == JsonValueKind.Array)
                {
                    // Write headers based on the first object's property names
                    JsonElement firstItem = root[0];
                    int colIndex = 0;
                    foreach (JsonProperty prop in firstItem.EnumerateObject())
                    {
                        sheet.Cells[0, colIndex].PutValue(prop.Name);
                        colIndex++;
                    }

                    // Write each array element as a row
                    int rowIndex = 1;
                    foreach (JsonElement item in root.EnumerateArray())
                    {
                        colIndex = 0;
                        foreach (JsonProperty prop in item.EnumerateObject())
                        {
                            string cellValue = prop.Value.ValueKind switch
                            {
                                JsonValueKind.String => prop.Value.GetString(),
                                JsonValueKind.Number => prop.Value.GetRawText(),
                                JsonValueKind.True => "True",
                                JsonValueKind.False => "False",
                                JsonValueKind.Null => null,
                                _ => prop.Value.GetRawText()
                            };
                            sheet.Cells[rowIndex, colIndex].PutValue(cellValue);
                            colIndex++;
                        }
                        rowIndex++;
                    }
                }
                else if (root.ValueKind == JsonValueKind.Object)
                {
                    // For a single object, write each property on a separate row (key/value)
                    int rowIndex = 0;
                    foreach (JsonProperty prop in root.EnumerateObject())
                    {
                        sheet.Cells[rowIndex, 0].PutValue(prop.Name);
                        string cellValue = prop.Value.ValueKind switch
                        {
                            JsonValueKind.String => prop.Value.GetString(),
                            JsonValueKind.Number => prop.Value.GetRawText(),
                            JsonValueKind.True => "True",
                            JsonValueKind.False => "False",
                            JsonValueKind.Null => null,
                            _ => prop.Value.GetRawText()
                        };
                        sheet.Cells[rowIndex, 1].PutValue(cellValue);
                        rowIndex++;
                    }
                }
                else
                {
                    // Unsupported JSON structure
                    sheet.Cells[0, 0].PutValue("Unsupported JSON structure.");
                }

                parsed = true;
            }
            catch (JsonException)
            {
                // If JSON parsing fails, write raw content to the first cell
                sheet.Cells[0, 0].PutValue(jsonContent);
            }

            // Auto-fit columns for better readability
            sheet.AutoFitColumns();

            // Save the workbook to an Excel file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Excel workbook created successfully at '{outputPath}'.");
        }
    }
}