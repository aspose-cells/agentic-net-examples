using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace ExcelXmlToJson
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file in XML format (Excel 2003 XML)
            string sourceXmlPath = "input.xml";

            // Path where the resulting JSON will be saved
            string jsonOutputPath = "output.json";

            // Load the XML‑based Excel file into a Workbook object
            Workbook workbook = new Workbook(sourceXmlPath);

            // Access the first worksheet (adjust if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Determine the used range of the worksheet
            int lastRow = worksheet.Cells.MaxDataRow;          // zero‑based index of the last row with data
            int lastColumn = worksheet.Cells.MaxDataColumn;    // zero‑based index of the last column with data

            // Create a range that covers all used cells (+1 because CreateRange expects count, not last index)
            Aspose.Cells.Range usedRange = worksheet.Cells.CreateRange(0, 0, lastRow + 1, lastColumn + 1);

            // Configure JSON export options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportEmptyCells = true,   // keep empty cells in the JSON output
                HasHeaderRow = true,       // treat the first row as header
                ExportNestedStructure = false
            };

            // Export the range to a JSON string using the provided rule
            string jsonResult = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);

            // Write the JSON string to the output file
            File.WriteAllText(jsonOutputPath, jsonResult);

            Console.WriteLine($"XML workbook has been converted to JSON and saved to '{jsonOutputPath}'.");
        }
    }
}