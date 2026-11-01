// Title: Create a case‑insensitive header‑to‑column index dictionary by iterating the first row with Aspose.Cells for .NET (C#)
// Common Searches: how to generate a dictionary of column indexes from Excel header row using Aspose.Cells C# | Aspose.Cells iterate first worksheet row to map header names to column numbers | C# read Excel headers and get column index with case‑insensitive lookup | build header lookup table for .xlsx file using Aspose.Cells API | retrieve column position of 'Amount' header in workbook with Aspose.Cells
// Tags: Aspose.Cells header index dictionary | case-insensitive column lookup in .NET | iterate worksheet first row cells | Excel .xlsx header mapping C# | column number extraction by header name

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads or creates an Excel workbook, iterates the first row to build a case‑insensitive Dictionary<string,int> that maps each non‑empty header to its column index, demonstrates fetching the index for the "Amount" column, and saves the workbook.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            Workbook workbook = null;

            try
            {
                if (File.Exists(inputPath))
                {
                    // Load existing workbook
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    // Create a new workbook with a sample header row
                    workbook = new Workbook();
                    Worksheet sheet = workbook.Worksheets[0];
                    var sheetCells = sheet.Cells;

                    string[] sampleHeaders = { "Id", "Name", "Amount", "Date" };
                    for (int i = 0; i < sampleHeaders.Length; i++)
                    {
                        sheetCells[0, i].PutValue(sampleHeaders[i]);
                    }
                }
            }
            catch (CellsException ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Ensure workbook was created/loaded
            if (workbook == null)
            {
                Console.WriteLine("Workbook is not initialized.");
                return;
            }

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Build a lookup dictionary for header names to column indexes
            var headerLookup = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            const int headerRowIndex = 0;
            int maxColumn = cells.MaxColumn;

            for (int col = 0; col <= maxColumn; col++)
            {
                string header = cells[headerRowIndex, col].StringValue?.Trim();
                if (string.IsNullOrEmpty(header))
                    continue;

                if (!headerLookup.ContainsKey(header))
                    headerLookup.Add(header, col);
            }

            // Example: retrieve column index for "Amount"
            if (headerLookup.TryGetValue("Amount", out int amountColumn))
            {
                Console.WriteLine($"Column index for 'Amount' is {amountColumn}");
            }
            else
            {
                Console.WriteLine("Header 'Amount' not found.");
            }

            // Optional: save the workbook
            try
            {
                workbook.Save("output.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
