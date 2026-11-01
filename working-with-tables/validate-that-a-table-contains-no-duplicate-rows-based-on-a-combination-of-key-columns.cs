// Title: Detect and Highlight Duplicate Rows in an Excel ListObject Using Aspose.Cells for .NET (C#)
// Common Searches: C# code to find duplicate rows in an Excel table with Aspose.Cells | how to highlight duplicate records in a ListObject using Aspose.Cells .NET | validate unique composite key (ID, Date) in Excel worksheet with Aspose.Cells | Aspose.Cells C# example for marking duplicate rows with custom style | detect duplicate rows based on multiple columns in an .xlsx file using Aspose.Cells
// Tags: duplicate row detection Aspose.Cells ListObject | highlight duplicate rows Excel C# | composite key validation Aspose.Cells | apply custom style to cells Aspose.Cells | hashset uniqueness check C# Excel

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsDuplicateValidator
{
    // The sample loads an Excel workbook, accesses the first ListObject table, defines "ID" and "Date" as key columns, and uses a HashSet to identify rows with identical composite keys. Duplicate rows are logged and then highlighted with a LightCoral background style before the workbook is saved to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output_with_duplicates_highlighted.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one table (ListObject)
                if (worksheet.ListObjects.Count == 0)
                {
                    Console.WriteLine("No tables found in the worksheet.");
                    return;
                }

                // Retrieve the first table
                ListObject table = worksheet.ListObjects[0];

                // Define key columns (as they appear in the table header)
                string[] keyColumnNames = { "ID", "Date" };

                // Map column names to zero‑based indices within the table
                Dictionary<string, int> columnNameToIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                for (int col = 0; col < table.ListColumns.Count; col++)
                {
                    columnNameToIndex[table.ListColumns[col].Name] = col;
                }

                // Verify that all key columns exist in the table
                foreach (string keyName in keyColumnNames)
                {
                    if (!columnNameToIndex.ContainsKey(keyName))
                    {
                        Console.WriteLine($"Key column \"{keyName}\" not found in the table.");
                        return;
                    }
                }

                // Track unique composite keys
                HashSet<string> uniqueKeys = new HashSet<string>();
                bool duplicateFound = false;

                // First pass: detect duplicates
                for (int row = 0; row < table.DataRange.RowCount; row++)
                {
                    List<string> keyParts = new List<string>();
                    foreach (string keyName in keyColumnNames)
                    {
                        int colIndex = columnNameToIndex[keyName];
                        object cellValue = table.DataRange[row, colIndex].Value;
                        keyParts.Add(cellValue?.ToString() ?? string.Empty);
                    }

                    string compositeKey = string.Join("||", keyParts);

                    if (!uniqueKeys.Add(compositeKey))
                    {
                        duplicateFound = true;
                        Console.WriteLine($"Duplicate row detected at table row {row + 2} (Excel row {table.DataRange.FirstRow + row + 2}).");
                    }
                }

                if (!duplicateFound)
                {
                    Console.WriteLine("No duplicate rows found based on the specified key columns.");
                    return;
                }

                // Prepare style for duplicate rows
                Style duplicateStyle = workbook.CreateStyle();
                duplicateStyle.ForegroundColor = Color.LightCoral;
                duplicateStyle.Pattern = BackgroundType.Solid;

                StyleFlag styleFlag = new StyleFlag { CellShading = true };

                // Second pass: apply style to duplicate rows
                uniqueKeys.Clear(); // reuse set
                for (int row = 0; row < table.DataRange.RowCount; row++)
                {
                    List<string> keyParts = new List<string>();
                    foreach (string keyName in keyColumnNames)
                    {
                        int colIndex = columnNameToIndex[keyName];
                        object cellValue = table.DataRange[row, colIndex].Value;
                        keyParts.Add(cellValue?.ToString() ?? string.Empty);
                    }

                    string compositeKey = string.Join("||", keyParts);

                    if (!uniqueKeys.Add(compositeKey))
                    {
                        // Apply style to each cell in the duplicate row
                        for (int col = 0; col < table.DataRange.ColumnCount; col++)
                        {
                            Cell cell = table.DataRange[row, col];
                            cell.SetStyle(duplicateStyle, styleFlag);
                        }
                    }
                }

                // Save the workbook with highlights
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved with duplicate rows highlighted: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
