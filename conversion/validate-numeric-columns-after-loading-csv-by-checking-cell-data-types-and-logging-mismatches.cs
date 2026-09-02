// Title: Check numeric columns in a CSV with Aspose.Cells for .NET and log type mismatches
// AI Prompts: Load a CSV file using Aspose.Cells, iterate through each column, and output the addresses of cells whose values are not numeric. | Write a .NET routine that validates numeric columns in a CSV, records non‑numeric entries together with their headers, and saves the workbook as an XLSX file.
// Common Searches: aspnet validate numeric data in CSV using Aspose.Cells | how to detect non‑numeric values in CSV columns with Aspose.Cells .NET | log cell type mismatches when converting CSV to Excel with Aspose.Cells
// Tags: numeric column validation Aspose.Cells | log non‑numeric cells .NET | CSV to XLSX data type checking Aspose.Cells | cell type verification Aspose.Cells CSV

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsCsvValidation
{
    // The program loads a CSV file with Aspose.Cells (enabling numeric and datetime conversion), scans each column for cells that are not numeric, logs mismatches with column headers and cell addresses, and finally saves the validated workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the CSV file to be validated
            string csvFilePath = "input.csv";

            // Load options for CSV import
            TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv)
            {
                Separator = ',',          // CSV delimiter
                ConvertNumericData = true, // Convert numeric strings to numeric values
                ConvertDateTimeData = true // Also convert date strings if present
            };

            // Load the CSV into a workbook
            Workbook workbook = new Workbook(csvFilePath, loadOptions);
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Determine the used range
            int maxRow = cells.MaxDataRow;      // zero‑based index of the last row with data
            int maxCol = cells.MaxDataColumn;   // zero‑based index of the last column with data

            // Assume first row contains headers
            for (int col = 0; col <= maxCol; col++)
            {
                // Header name for reporting
                string header = cells[0, col].StringValue;

                // Scan data rows for non‑numeric values
                for (int row = 1; row <= maxRow; row++)
                {
                    Cell cell = cells[row, col];

                    // Skip empty cells
                    if (cell.Type == CellValueType.IsNull)
                        continue;

                    // If the cell is not numeric (int, double, or datetime) log the mismatch
                    if (!cell.IsNumericValue)
                    {
                        Console.WriteLine($"Mismatch in column '{header}' (Index {col + 1}) at cell {cell.Name}: " +
                                          $"Value=\"{cell.StringValue}\" is not numeric.");
                    }
                }
            }

            // Optionally, save the workbook after validation (e.g., to Excel format)
            workbook.Save("validated_output.xlsx", SaveFormat.Xlsx);
        }
    }
}
