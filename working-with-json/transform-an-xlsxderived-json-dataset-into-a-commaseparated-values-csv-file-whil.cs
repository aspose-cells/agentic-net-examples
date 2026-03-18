using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonToCsv
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file (optional, if you already have JSON you can skip loading XLSX)
            string xlsxPath = "source.xlsx";

            // Path for the output CSV file
            string csvPath = "output.csv";

            // -----------------------------------------------------------------
            // Step 1: Load the XLSX workbook (if you need to extract JSON from it)
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook(xlsxPath);

            // -----------------------------------------------------------------
            // Step 2: Export the worksheet range to JSON (preserving structure)
            // -----------------------------------------------------------------
            // Export the used range of the first worksheet to JSON
            Worksheet sheet = workbook.Worksheets[0];
            int firstRow = sheet.Cells.MaxDisplayRange.FirstRow;
            int firstColumn = sheet.Cells.MaxDisplayRange.FirstColumn;
            int totalRows = sheet.Cells.MaxDisplayRange.RowCount;
            int totalColumns = sheet.Cells.MaxDisplayRange.ColumnCount;

            // Create JSON save options to keep Excel-like structure
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ToExcelStruct = true,
                ExportEmptyCells = true,
                HasHeaderRow = true
            };

            // Export the range to a JSON string
            string json = JsonUtility.ExportRangeToJson(
                sheet.Cells.CreateRange(firstRow, firstColumn, totalRows, totalColumns),
                jsonOptions);

            // -----------------------------------------------------------------
            // Step 3: Create a new workbook and import the JSON data back
            // -----------------------------------------------------------------
            Workbook jsonWorkbook = new Workbook();
            Worksheet jsonSheet = jsonWorkbook.Worksheets[0];

            // Use JsonLayoutOptions to treat arrays as tables
            JsonLayoutOptions layoutOptions = new JsonLayoutOptions
            {
                ArrayAsTable = true
            };

            // Import JSON into the worksheet starting at cell A1 (row 0, column 0)
            JsonUtility.ImportData(json, jsonSheet.Cells, 0, 0, layoutOptions);

            // -----------------------------------------------------------------
            // Step 4: Save the workbook as CSV, preserving column alignment
            // -----------------------------------------------------------------
            jsonWorkbook.Save(csvPath, SaveFormat.Csv);

            Console.WriteLine($"JSON data from '{xlsxPath}' has been converted to CSV at '{csvPath}'.");
        }
    }
}