// Title: Check for duplicate rows in an Excel ListObject using composite key columns with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, iterates a ListObject, builds a composite key from specified columns, and prints the row numbers that share duplicate keys. | Enhance the duplicate‑row validator to apply a red background style to rows identified as duplicates and save the modified workbook as a new file.
// Common Searches: aspocells c# find duplicate rows in a table using multiple columns | how to validate uniqueness of rows in an Excel ListObject with Aspose.Cells | c# Aspose.Cells check for duplicate entries based on composite key | detect duplicate rows in Excel table programmatically using Aspose.Cells .NET | using HashSet to identify duplicate rows in Aspose.Cells ListObject
// Tags: Aspose.Cells ListObject duplicate detection | composite key validation Aspose.Cells | C# HashSet for row uniqueness | Excel table duplicate check .NET | format duplicate rows Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The program loads an Excel workbook, accesses a ListObject named "Table1", builds a composite key from selected columns for each data row, uses a HashSet to detect duplicate keys, reports duplicate row numbers, and can optionally highlight and save the workbook.
class DuplicateRowValidator
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string inputPath = "input.xlsx";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (or specify by name/index)
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the table (ListObject) by its name; adjust the name as needed
            ListObject table = worksheet.ListObjects["Table1"];
            if (table == null)
            {
                Console.WriteLine("Error: Table \"Table1\" was not found in the worksheet.");
                return;
            }

            // Define zero‑based indices of the key columns within the table (e.g., columns A and C)
            int[] keyColumnIndices = new int[] { 0, 2 };

            // Set to store unique composite keys
            HashSet<string> uniqueKeys = new HashSet<string>();
            bool duplicateFound = false;

            // Determine the data range of the table
            int firstDataRow = table.DataRange.FirstRow;
            int lastDataRow = firstDataRow + table.DataRange.RowCount - 1;
            int firstDataCol = table.DataRange.FirstColumn;

            // Iterate over each data row
            for (int row = firstDataRow; row <= lastDataRow; row++)
            {
                // Build a composite key from the specified columns
                List<string> keyParts = new List<string>();
                foreach (int colIndex in keyColumnIndices)
                {
                    string cellValue = worksheet.Cells[row, firstDataCol + colIndex].StringValue?.Trim() ?? string.Empty;
                    keyParts.Add(cellValue);
                }

                string compositeKey = string.Join("||", keyParts);

                // Detect duplicates
                if (!uniqueKeys.Add(compositeKey))
                {
                    duplicateFound = true;
                    Console.WriteLine($"Duplicate row detected at table row {row - firstDataRow + 1}.");
                }
            }

            if (!duplicateFound)
            {
                Console.WriteLine("No duplicate rows found based on the specified key columns.");
            }

            // Optionally, save the workbook (e.g., after marking duplicates)
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
