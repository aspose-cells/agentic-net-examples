// Title: Validate numeric columns in a CSV with Aspose.Cells for .NET (C#)
// Description: Load a CSV into an Aspose.Cells workbook, enable numeric conversion, iterate through the data rows, use IsNumericValue to detect non‑numeric cells, log each mismatch with row, column and cell name, and optionally save the unchanged workbook. Ideal for data‑quality checks after CSV import.
// Keywords: Aspose.Cells | C# CSV import | numeric validation | IsNumericValue | TxtLoadOptions | ConvertNumericData | cell type checking | log non‑numeric cells | Excel workbook | data quality
// Common Searches: Aspose.Cells validate numeric values in CSV | C# check non‑numeric cells after CSV import | How to log mismatched data types with Aspose.Cells | Detect text in numeric columns using Aspose.Cells | CSV data validation example in .NET
// Developer Intent: Ensure every cell that should contain a number after importing a CSV is actually numeric and capture any deviations for review.
// Use Cases: Generate a list of cell addresses that contain text where numbers are expected, to clean source data before calculations. | Create an audit report of data‑type errors in CSV files imported into Excel worksheets. | Prevent runtime errors in downstream analytics by flagging non‑numeric entries immediately after import.
// AI Prompts: Write C# code with Aspose.Cells that validates numeric columns in an imported CSV and stores mismatched cell references in a List<string>. | Show how to modify the validation loop to skip multiple header rows and validate only columns whose header matches a given pattern. | Explain how to configure TxtLoadOptions so that empty strings are treated as null and included in the mismatch log.

using System;
using System.IO;
using Aspose.Cells;

namespace CsvNumericValidationDemo
{
    // Load a CSV into an Aspose.Cells workbook, enable numeric conversion, iterate through the data rows, use IsNumericValue to detect non‑numeric cells, log each mismatch with row, column and cell name, and optionally save the unchanged workbook. Ideal for data‑quality checks after CSV import.
    class Program
    {
        static void Main()
        {
            // Path to the CSV file to be validated
            string csvFilePath = "input.csv";

            // Create a new workbook (empty workbook)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Configure CSV load options
            TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv)
            {
                Separator = ',',            // CSV delimiter
                ConvertNumericData = true, // Convert numeric strings to numeric values
                ConvertDateTimeData = true // Convert date strings if needed
            };

            // Import CSV data starting at cell A1 (row 0, column 0)
            cells.ImportCSV(csvFilePath, loadOptions, 0, 0);

            // Determine the used range
            int maxRow = cells.MaxDataRow;       // zero‑based index of the last row with data
            int maxColumn = cells.MaxDataColumn; // zero‑based index of the last column with data

            // Assume first row (row 0) contains headers, start validation from row 1
            for (int col = 0; col <= maxColumn; col++)
            {
                bool columnHasNumericHeader = true; // optional: you could check header type here

                for (int row = 1; row <= maxRow; row++)
                {
                    Cell cell = cells[row, col];

                    // If the cell is empty, skip it
                    if (cell.Type == CellValueType.IsNull)
                        continue;

                    // Check whether the cell value is numeric (int, double, or datetime)
                    if (!cell.IsNumericValue)
                    {
                        // Log mismatch: column index, row index (1‑based for readability), and cell content
                        Console.WriteLine($"Mismatch found at Row {row + 1}, Column {col + 1} (Cell {cell.Name}): " +
                                          $"Value \"{cell.StringValue}\" is not numeric.");
                    }
                }
            }

            // Save the workbook (optional – the workbook is unchanged)
            workbook.Save("validated_output.xlsx", SaveFormat.Xlsx);
        }
    }
}
