using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsCsvImportExample
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the CSV file (replace with your actual file path)
                string csvPath = "data.csv";

                // Verify that the CSV file exists to avoid FileNotFoundException
                if (!File.Exists(csvPath))
                {
                    Console.WriteLine($"CSV file not found: {csvPath}");
                    return;
                }

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Configure CSV load options to automatically detect numeric and date values
                TxtLoadOptions loadOptions = new TxtLoadOptions
                {
                    Separator = ',',               // CSV delimiter
                    ConvertNumericData = true,     // Convert numeric strings to numbers
                    ConvertDateTimeData = true,    // Convert date strings to DateTime
                    ParsingFormulaOnOpen = false   // Not needed for plain CSV
                };

                // Import CSV data starting at cell A1 (row 0, column 0)
                cells.ImportCSV(csvPath, loadOptions, 0, 0);

                // Determine the range that contains the imported data
                int lastRow = cells.MaxDataRow;
                int lastColumn = cells.MaxDataColumn;

                // Create a table (ListObject) from the imported range
                // The 'true' parameter indicates that the first row contains column headers
                int tableIndex = worksheet.ListObjects.Add(0, 0, lastRow + 1, lastColumn + 1, true);
                ListObject table = worksheet.ListObjects[tableIndex];
                table.DisplayName = "ImportedCsvTable";

                // Save the workbook with the new table
                string outputPath = "ImportedCsvWithTable.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}