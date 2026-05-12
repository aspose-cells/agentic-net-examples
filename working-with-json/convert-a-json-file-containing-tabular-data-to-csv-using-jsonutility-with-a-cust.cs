using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonToCsv
{
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input JSON and output CSV
            string jsonFilePath = "input.json";
            string csvFilePath = "output.csv";

            // Custom delimiter for CSV (e.g., pipe character)
            string delimiter = "|";

            // ------------------- Create Workbook -------------------
            Workbook workbook = new Workbook();                     // create rule
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // ------------------- Load JSON content -------------------
            string jsonContent = File.ReadAllText(jsonFilePath);

            // ------------------- Import JSON to cells -------------------
            JsonLayoutOptions layoutOptions = new JsonLayoutOptions
            {
                ArrayAsTable = true,          // treat JSON array as a table
                ConvertNumericOrDate = true   // optional: convert numbers/dates
            };
            JsonUtility.ImportData(jsonContent, cells, 0, 0, layoutOptions); // import rule

            // ------------------- Determine used range -------------------
            int maxRow = cells.MaxDataRow;       // last row with data
            int maxCol = cells.MaxDataColumn;    // last column with data

            // ------------------- Build CSV content -------------------
            using (StreamWriter writer = new StreamWriter(csvFilePath))
            {
                for (int row = 0; row <= maxRow; row++)
                {
                    string[] rowValues = new string[maxCol + 1];
                    for (int col = 0; col <= maxCol; col++)
                    {
                        var cell = cells[row, col];
                        string cellText = cell.IsFormula ? cell.Formula : cell.StringValue;

                        // Escape delimiter and quotes if needed
                        if (cellText.Contains(delimiter) || cellText.Contains("\""))
                        {
                            cellText = $"\"{cellText.Replace("\"", "\"\"")}\"";
                        }

                        rowValues[col] = cellText;
                    }
                    string line = string.Join(delimiter, rowValues);
                    writer.WriteLine(line);
                }
            }

            // ------------------- Optional: Save workbook as Excel (lifecycle rule) -------------------
            workbook.Save("intermediate.xlsx"); // save rule (optional, demonstrates lifecycle)

            Console.WriteLine($"JSON data has been converted to CSV with delimiter '{delimiter}'.");
        }
    }
}