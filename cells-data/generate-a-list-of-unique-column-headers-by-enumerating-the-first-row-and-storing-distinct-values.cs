// Title: How to list distinct column headers from the first row of an Excel file using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells read first row and get unique header names | extract distinct column names from Excel worksheet using Aspose.Cells .NET | enumerate header row in Excel and remove duplicates with Aspose.Cells | get list of non‑empty column headers from first row in C# Aspose.Cells | handle password‑protected workbook while extracting unique headers Aspose.Cells
// Tags: Aspose.Cells distinct header extraction | C# read first row Excel | HashSet case‑insensitive header collection | Workbook load error handling Aspose.Cells | MaxDataColumn usage Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook with Aspose.Cells, safely handles missing or password‑protected files, accesses the first worksheet, determines the last populated column in the first row, and uses a case‑insensitive HashSet to collect non‑empty, trimmed header strings, finally printing the distinct column headers.
    class Program
    {
        static void Main()
        {
            const string filePath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            Workbook workbook = null;
            try
            {
                // Attempt to load the workbook (no password supplied)
                workbook = new Workbook(filePath);
            }
            catch (CellsException ex)
            {
                // Handle password‑protected files or other loading issues
                if (!string.IsNullOrEmpty(ex.Message) &&
                    ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Console.WriteLine("The workbook is password protected and cannot be opened without a password.");
                }
                else
                {
                    Console.WriteLine($"Error loading workbook: {ex.Message}");
                }
                return;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error loading workbook: {ex.Message}");
                return;
            }

            try
            {
                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Determine the last column that contains data in the first row
                int lastColumn = sheet.Cells.MaxDataColumn;

                // Store unique header values (case‑insensitive)
                HashSet<string> uniqueHeaders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                // Enumerate the first row and collect distinct values
                for (int col = 0; col <= lastColumn; col++)
                {
                    Cell cell = sheet.Cells[0, col];
                    string header = cell.StringValue?.Trim();

                    if (!string.IsNullOrEmpty(header))
                    {
                        uniqueHeaders.Add(header);
                    }
                }

                // Output the unique column headers
                Console.WriteLine("Unique column headers:");
                foreach (string header in uniqueHeaders)
                {
                    Console.WriteLine(header);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}
