// Title: Create a case‑insensitive header‑to‑column index map with Aspose.Cells for .NET (C#)
// Description: This example shows how to load or create an Excel workbook, write sample headers, and iterate the first row to build a Dictionary<string,int> that links each non‑empty header (trimmed and case‑insensitive) to its zero‑based column index. The mapping is printed to the console and the workbook is saved.
// Keywords: Aspose.Cells header index mapping | C# Excel column lookup | case insensitive header dictionary | iterate first row Aspose.Cells | Excel header to column number .NET | dynamic column reference C# | Aspose.Cells sample code
// Common Searches: how to map Excel headers to column numbers using Aspose.Cells C# | retrieve column index by header name Aspose.Cells .NET | create dictionary of sheet headers Aspose.Cells | skip empty header cells Aspose.Cells example | case‑insensitive header lookup in Excel with C#
// Developer Intent: Generate a fast, case‑insensitive lookup that returns the column index for any header in the first worksheet row.
// Use Cases: Read rows by referencing cells through header names instead of hard‑coded indexes. | Validate required columns before processing and raise errors for missing headers. | Construct dynamic SQL statements or data‑transfer objects by translating header names to positions at runtime.
// AI Prompts: Write C# code using Aspose.Cells that builds a case‑insensitive dictionary of header names to column indices, ignoring blank cells and handling duplicate headers. | Extend the example to combine multiple header rows into a single lookup dictionary, preserving hierarchy and allowing composite keys.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHeaderMapping
{
    // This example shows how to load or create an Excel workbook, write sample headers, and iterate the first row to build a Dictionary<string,int> that links each non‑empty header (trimmed and case‑insensitive) to its zero‑based column index. The mapping is printed to the console and the workbook is saved.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook(); // replace with new Workbook("input.xlsx") to load
                Worksheet worksheet = workbook.Worksheets[0];

                // Sample data: add header row for demonstration
                worksheet.Cells["A1"].PutValue("ID");
                worksheet.Cells["B1"].PutValue("Name");
                worksheet.Cells["C1"].PutValue("Price");
                worksheet.Cells["D1"].PutValue("Quantity");

                // Dictionary to hold header name -> column index mapping
                Dictionary<string, int> headerToIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                // Assume the header row is the first row (index 0)
                int headerRowIndex = 0;
                Cells cells = worksheet.Cells;

                // Determine the last column with data in the worksheet (fallback if specific row method unavailable)
                int lastColumn = cells.MaxColumn;

                // Iterate through each column in the header row
                for (int col = 0; col <= lastColumn; col++)
                {
                    string? header = cells[headerRowIndex, col].StringValue;

                    // Skip empty or whitespace header cells
                    if (string.IsNullOrWhiteSpace(header))
                        continue;

                    header = header.Trim();

                    // Add to dictionary (if duplicate header, later one overwrites)
                    headerToIndex[header] = col;
                }

                // Example usage: print the mapping
                Console.WriteLine("Header to Column Index Mapping:");
                foreach (var kvp in headerToIndex)
                {
                    Console.WriteLine($"Header \"{kvp.Key}\" => Column Index {kvp.Value}");
                }

                // Save the workbook if needed
                string outputPath = "HeaderMappingDemo.xlsx";

                // Ensure the directory exists before saving
                string? directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
