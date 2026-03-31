using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace JsonToExcelConverter
{
    // Classes that represent the expected JSON structure
    public class WorkbookJson
    {
        public List<SheetJson> Sheets { get; set; }
    }

    public class SheetJson
    {
        public string Name { get; set; }
        public List<List<JsonElement>> Data { get; set; }
    }

    public class Program
    {
        // Entry point: args[0] = input JSON file path, args[1] = output Excel file path
        public static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: JsonToExcelConverter <inputJsonPath> <outputExcelPath>");
                return;
            }

            string jsonPath = args[0];
            string excelPath = args[1];

            // Read and deserialize the JSON representation
            string jsonContent = File.ReadAllText(jsonPath);
            WorkbookJson workbookData = JsonSerializer.Deserialize<WorkbookJson>(jsonContent);

            // Create a new workbook (default format is Xlsx)
            Workbook workbook = new Workbook();

            // Process each sheet defined in the JSON
            for (int sheetIndex = 0; sheetIndex < workbookData.Sheets.Count; sheetIndex++)
            {
                SheetJson sheetInfo = workbookData.Sheets[sheetIndex];
                Worksheet worksheet;

                if (sheetIndex == 0)
                {
                    // Use the default first worksheet
                    worksheet = workbook.Worksheets[0];
                }
                else
                {
                    // Add a new worksheet for subsequent sheets
                    int newIndex = workbook.Worksheets.Add();
                    worksheet = workbook.Worksheets[newIndex];
                }

                // Set the worksheet name if provided
                if (!string.IsNullOrEmpty(sheetInfo.Name))
                {
                    worksheet.Name = sheetInfo.Name;
                }

                // Populate cells with the data matrix
                for (int row = 0; row < sheetInfo.Data.Count; row++)
                {
                    List<JsonElement> rowData = sheetInfo.Data[row];
                    for (int col = 0; col < rowData.Count; col++)
                    {
                        JsonElement cellValue = rowData[col];
                        // Determine the appropriate .NET type based on JSON value kind
                        object value = cellValue.ValueKind switch
                        {
                            JsonValueKind.String => cellValue.GetString(),
                            JsonValueKind.Number => cellValue.TryGetInt64(out long l) ? (object)l :
                                                    cellValue.TryGetDouble(out double d) ? d : null,
                            JsonValueKind.True => true,
                            JsonValueKind.False => false,
                            JsonValueKind.Null => null,
                            _ => cellValue.GetRawText()
                        };

                        // Put the value into the cell (null will clear the cell)
                        worksheet.Cells[row, col].PutValue(value);
                    }
                }
            }

            // Save the populated workbook to the specified file in XLSX format
            workbook.Save(excelPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook successfully created at: {excelPath}");
        }
    }
}