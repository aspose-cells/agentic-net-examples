using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonImportDemo
{
    public class Program
    {
        public static void Main()
        {
            // Sample JSON string representing an array of objects
            string json = "[{"
                        + "\"Name\":\"John\","
                        + "\"Age\":30"
                        + "},"
                        + "{"
                        + "\"Name\":\"Alice\","
                        + "\"Age\":25"
                        + "}]";

            // Create a new workbook (in-memory Excel file)
            Workbook workbook = new Workbook();

            // Access the first worksheet where data will be placed
            Worksheet worksheet = workbook.Worksheets[0];

            // Set JSON layout options – treat the JSON array as a table
            JsonLayoutOptions options = new JsonLayoutOptions();
            options.ArrayAsTable = true;   // each array element becomes a row

            // Import the JSON data into the worksheet starting at cell A1 (row 0, column 0)
            JsonUtility.ImportData(json, worksheet.Cells, 0, 0, options);

            // Save the workbook in XLSX format
            workbook.Save("JsonImported.xlsx");
        }
    }
}