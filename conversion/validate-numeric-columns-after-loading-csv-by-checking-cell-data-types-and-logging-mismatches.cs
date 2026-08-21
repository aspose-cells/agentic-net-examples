// Title: C# – Validate Numeric Columns in a CSV with Aspose.Cells and Log Type Mismatches
// Description: Loads a CSV using TxtLoadOptions (numeric and date conversion enabled), detects columns that contain numeric data, checks each cell in those columns for the correct data type, writes mismatched entries to the console, and saves the workbook as an XLSX file for review.
// Keywords: Aspose.Cells CSV numeric validation | C# validate numeric columns | cell data type mismatch logging | TxtLoadOptions ConvertNumericData | detect non‑numeric values in Excel worksheet | CSV to XLSX conversion Aspose.Cells | data quality check Aspose.Cells
// Common Searches: how to validate numeric columns in a CSV using Aspose.Cells .NET | log cells that are not numeric after loading CSV with Aspose.Cells | detect numeric columns and report type errors in C# | save validated CSV as Excel with Aspose.Cells | Aspose.Cells convert numeric strings automatically
// Developer Intent: Identify columns that should contain numbers after loading a CSV and report any cells that hold non‑numeric values.
// Use Cases: Ensure data integrity before performing calculations or imports. | Generate a report of rows where text appears in numeric fields. | Automate quality checks for CSV files received from external systems. | Create an Excel file that highlights mismatched cells for manual review.
// AI Prompts: Generate C# code with Aspose.Cells that loads a CSV, finds numeric columns, and prints addresses of cells that are not numeric. | Provide a reusable method that returns a list of cell names where expected numeric values are stored as text after CSV import. | Explain how TxtLoadOptions settings affect numeric and date conversion, then show how to validate cell data types in the resulting worksheet.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsNumericValidation
{
    // Loads a CSV using TxtLoadOptions (numeric and date conversion enabled), detects columns that contain numeric data, checks each cell in those columns for the correct data type, writes mismatched entries to the console, and saves the workbook as an XLSX file for review.
    class Program
    {
        static void Main()
        {
            // Path to the CSV file to be validated
            string csvPath = "data.csv";

            // Load options: enable automatic conversion of numeric strings to numeric values
            TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv)
            {
                Separator = ',',               // CSV delimiter
                ConvertNumericData = true,     // Convert numeric strings during load
                ConvertDateTimeData = true     // Convert date strings if present
            };

            // Load the CSV into a workbook using the load options
            Workbook workbook = new Workbook(csvPath, loadOptions);
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Determine the used range
            int maxRow = cells.MaxDataRow;
            int maxColumn = cells.MaxDataColumn;

            // Assume the first row contains headers; start validation from the second row (index 1)
            int dataStartRow = 1;

            // Identify columns that contain at least one numeric value (excluding header)
            bool[] isNumericColumn = new bool[maxColumn + 1];
            for (int col = 0; col <= maxColumn; col++)
            {
                for (int row = dataStartRow; row <= maxRow; row++)
                {
                    Cell cell = cells[row, col];
                    if (cell != null && cell.IsNumericValue)
                    {
                        isNumericColumn[col] = true;
                        break;
                    }
                }
            }

            // Validate each cell in identified numeric columns
            for (int col = 0; col <= maxColumn; col++)
            {
                if (!isNumericColumn[col]) continue; // Skip non‑numeric columns

                for (int row = dataStartRow; row <= maxRow; row++)
                {
                    Cell cell = cells[row, col];
                    if (cell == null) continue;

                    // If the cell is not numeric, log the mismatch
                    if (!cell.IsNumericValue)
                    {
                        Console.WriteLine($"Mismatch at {cell.Name}: Expected numeric, found '{cell.StringValue}'.");
                    }
                }
            }

            // Optionally save the workbook after validation (e.g., to review conversions)
            workbook.Save("validated_output.xlsx", SaveFormat.Xlsx);
        }
    }
}
