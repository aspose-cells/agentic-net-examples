using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

class Program
{
    static void Main()
    {
        try
        {
            // Verify that the CSV file exists to avoid FileNotFoundException
            string csvFilePath = "data.csv";
            if (!File.Exists(csvFilePath))
            {
                Console.WriteLine($"CSV file not found: {csvFilePath}");
                return;
            }

            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Import CSV data starting at cell A1 (comma delimiter, auto‑detect numeric values)
            cells.ImportCSV(csvFilePath, ",", true, 0, 0);

            // Determine the used range after import
            int lastRow = cells.MaxDataRow;
            int lastColumn = cells.MaxDataColumn;

            // Add a ListObject (Excel table) that covers the imported data
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
            int listObjectIndex = worksheet.ListObjects.Add(0, 0, lastRow + 1, lastColumn + 1, true);
            ListObject listObject = worksheet.ListObjects[listObjectIndex];

            // Set a friendly name for the table
            listObject.DisplayName = "CsvQueryTable";

            // Auto‑fit columns for better readability
            worksheet.AutoFitColumns();

            // Save the workbook
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}