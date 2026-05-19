using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsXlsbToJson
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSB file
            string xlsbPath = "input.xlsb";

            // Path where the generated JSON will be saved
            string jsonOutputPath = "output.json";

            // ---------- Load the XLSB workbook ----------
            // Create LoadOptions for XLSB format and ensure formulas are parsed on open
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsb)
            {
                ParsingFormulaOnOpen = true
            };

            // Load the workbook using the provided constructor (Workbook(string, LoadOptions))
            Workbook workbook = new Workbook(xlsbPath, loadOptions);

            // ---------- Calculate all formulas ----------
            // This uses the Workbook.CalculateFormula() method to evaluate every formula in the workbook
            workbook.CalculateFormula();

            // ---------- Convert workbook data to JSON ----------
            // We'll represent each worksheet as an array of rows, where each row is an array of cell values
            var workbookData = new Dictionary<string, object>();

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                var sheetData = new List<List<object>>();

                // Determine the used range
                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    var rowData = new List<object>();
                    for (int col = 0; col <= maxCol; col++)
                    {
                        // Retrieve the cell value; if the cell is empty, store null
                        object value = sheet.Cells[row, col].Value;
                        rowData.Add(value);
                    }
                    sheetData.Add(rowData);
                }

                workbookData[sheet.Name] = sheetData;
            }

            // Serialize the dictionary to a formatted JSON string
            string json = JsonSerializer.Serialize(workbookData, new JsonSerializerOptions { WriteIndented = true });

            // Write the JSON string to the output file
            File.WriteAllText(jsonOutputPath, json);

            Console.WriteLine($"Workbook data has been exported to JSON file: {jsonOutputPath}");
        }
    }
}