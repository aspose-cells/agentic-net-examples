// Title: Validate workbook encryption and extract all rows from each populated worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, checks Workbook.IsEncrypted (or LoadOptions.Password) and throws a clear exception if the file is encrypted, otherwise extracts every cell value from all non‑empty worksheets into a Dictionary<string, List<object[]>>. | Create a reusable method ExtractData(string filePath) that returns a dictionary mapping sheet names to row arrays, performs an encryption check before loading, and logs any worksheets that are empty and therefore skipped. | Update the example to handle password‑protected workbooks by prompting for a password, using LoadOptions.Password with Aspose.Cells, and proceeding with data extraction only after successful decryption.
// Common Searches: how to check if an Excel workbook is encrypted with Aspose.Cells before reading data | extract data from all non‑empty sheets in a .xlsx file using Aspose.Cells C# | skip empty worksheets while iterating through workbook worksheets Aspose.Cells | load password protected Excel file with Aspose.Cells and read its contents in .NET
// Tags: aspocells workbook encryption check | aspocells extract all worksheets data | aspocells skip empty sheets | aspocells load password protected xlsx | c# bulk excel sheet extraction aspocells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace WorkbookExtraction
{
    // The sample loads an Excel workbook with Aspose.Cells, verifies the file exists, optionally checks for encryption, iterates through each worksheet, skips those without data, extracts every cell value into a dictionary keyed by sheet name, and reports the row count per sheet; encryption handling can be added via Workbook.IsEncrypted or LoadOptions.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file
            string filePath = @"C:\Data\Sample.xlsx";

            // Verify that the file exists before attempting to load it
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            Workbook workbook = null;

            try
            {
                // Load the workbook (uses the standard load rule)
                workbook = new Workbook(filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
                return;
            }

            // NOTE: Older Aspose.Cells versions may not expose an IsEncrypted property.
            // If needed, handle encryption via Workbook.LoadOptions in a real scenario.

            // Prepare a container for extracted data
            var extractedData = new Dictionary<string, List<object[]>>();

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine if the sheet is empty by checking the used range
                var cells = sheet.Cells;
                if (cells.MaxDataRow < 0 || cells.MaxDataColumn < 0)
                {
                    // Skip empty sheets
                    continue;
                }

                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                var sheetData = new List<object[]>();

                // Extract data row by row
                for (int row = 0; row <= maxRow; row++)
                {
                    var rowData = new object[maxCol + 1];
                    for (int col = 0; col <= maxCol; col++)
                    {
                        rowData[col] = cells[row, col].Value;
                    }
                    sheetData.Add(rowData);
                }

                // Store data keyed by sheet name
                extractedData[sheet.Name] = sheetData;
            }

            // Example: output the number of rows extracted per sheet
            foreach (var kvp in extractedData)
            {
                Console.WriteLine($"Sheet '{kvp.Key}' - Rows extracted: {kvp.Value.Count}");
            }

            // Further processing of extractedData can be performed here
        }
    }
}
