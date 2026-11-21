// Title: Identify duplicate rows in an Excel ListObject using Aspose.Cells for C# with composite key columns
// Common Searches: C# code to find duplicate rows in an Excel table using Aspose.Cells | how to validate unique CustomerID and OrderDate combinations in a ListObject with Aspose.Cells | using HashSet to detect duplicate records in an Aspose.Cells worksheet | Aspose.Cells example for checking duplicate entries in a table via C# | detect duplicate rows based on multiple columns in Excel through Aspose.Cells API
// Tags: Aspose.Cells ListObject duplicate row detection | C# composite key uniqueness check in Excel tables | HashSet based duplicate detection with Aspose.Cells | mapping header names to column indices Aspose.Cells | Excel table row uniqueness validation C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsDuplicateValidator
{
    // The sample loads an Excel workbook, accesses the ListObject named "MyTable", maps the header names "CustomerID" and "OrderDate" to their column indices, builds a composite key for each data row, and uses a case‑insensitive HashSet to flag any duplicate rows, reporting the table and Excel row numbers where duplicates occur.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (adjust index or name as needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Retrieve the table (ListObject) named "MyTable"
                ListObject table = sheet.ListObjects["MyTable"];
                if (table == null)
                {
                    Console.WriteLine("Table \"MyTable\" not found in the worksheet.");
                    return;
                }

                // Define the key columns by their names (as they appear in the table header)
                string[] keyColumnNames = { "CustomerID", "OrderDate" };

                // Map column names to their zero‑based indices within the table
                int[] keyColumnIndices = new int[keyColumnNames.Length];
                for (int i = 0; i < keyColumnNames.Length; i++)
                {
                    int colIndex = -1;
                    for (int c = 0; c < table.ListColumns.Count; c++)
                    {
                        if (string.Equals(table.ListColumns[c].Name, keyColumnNames[i],
                                          StringComparison.OrdinalIgnoreCase))
                        {
                            colIndex = c;
                            break;
                        }
                    }

                    if (colIndex == -1)
                    {
                        Console.WriteLine($"Key column \"{keyColumnNames[i]}\" not found in the table.");
                        return;
                    }

                    keyColumnIndices[i] = colIndex;
                }

                // Use a HashSet to track unique key combinations
                HashSet<string> seenKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                bool duplicateFound = false;

                // Iterate over data rows (DataRange contains only data rows, zero‑based)
                for (int row = 0; row < table.DataRange.RowCount; row++)
                {
                    List<string> keyParts = new List<string>();
                    foreach (int colIdx in keyColumnIndices)
                    {
                        // Retrieve cell value from the data range
                        object cellValue = table.DataRange[row, colIdx].Value;
                        keyParts.Add(cellValue?.ToString() ?? string.Empty);
                    }

                    string compositeKey = string.Join("||", keyParts);

                    if (!seenKeys.Add(compositeKey))
                    {
                        duplicateFound = true;
                        int excelRowNumber = table.DataRange.FirstRow + row + 1; // +1 for Excel's 1‑based rows
                        Console.WriteLine($"Duplicate row found at table row {row + 2} (Excel row {excelRowNumber}).");
                    }
                }

                if (!duplicateFound)
                {
                    Console.WriteLine("No duplicate rows found based on the specified key columns.");
                }

                // Optionally, save the workbook if you made any changes
                // workbook.Save("output.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
