using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsValidationDemo
{
    class Program
    {
        static void Main()
        {
            // -------------------- Create --------------------
            // Create a new empty workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // -------------------- Load --------------------
            // Sample CSV data (first row = headers)
            string csvData = "ID,Value,Date\n" +
                             "1,123.45,2023-01-01\n" +
                             "2,abc,2023-01-02\n" +
                             "3,67.89,invalid";

            // Convert CSV string to a memory stream
            byte[] csvBytes = System.Text.Encoding.UTF8.GetBytes(csvData);
            using (MemoryStream csvStream = new MemoryStream(csvBytes))
            {
                // Configure load options to convert numeric strings automatically
                TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv)
                {
                    Separator = ',',
                    ConvertNumericData = true,
                    ConvertDateTimeData = true
                };

                // Import CSV data into the worksheet starting at cell A1 (row 0, column 0)
                cells.ImportCSV(csvStream, loadOptions, 0, 0);
            }

            // -------------------- Validate Numeric Columns --------------------
            // Determine which columns are expected to be numeric based on header names
            // (In this example, columns "ID" and "Value" should be numeric)
            int totalRows = cells.MaxDataRow + 1;      // includes header row
            int totalCols = cells.MaxDataColumn + 1;

            // Identify numeric columns
            bool[] isNumericColumn = new bool[totalCols];
            for (int col = 0; col < totalCols; col++)
            {
                string header = cells[0, col].StringValue?.Trim().ToUpperInvariant();
                if (header == "ID" || header == "VALUE")
                {
                    isNumericColumn[col] = true;
                }
            }

            // Scan data rows and log mismatches
            for (int row = 1; row < totalRows; row++) // start after header
            {
                for (int col = 0; col < totalCols; col++)
                {
                    if (!isNumericColumn[col]) continue; // skip non‑numeric columns

                    Cell cell = cells[row, col];
                    if (!cell.IsNumericValue)
                    {
                        string columnName = cells[0, col].StringValue;
                        Console.WriteLine($"Mismatch at {cell.Name}: Expected numeric in column '{columnName}', but found '{cell.StringValue}'.");
                    }
                }
            }

            // -------------------- Save --------------------
            // Save the workbook to verify the import (optional)
            workbook.Save("ValidatedOutput.xlsx", SaveFormat.Xlsx);
        }
    }
}