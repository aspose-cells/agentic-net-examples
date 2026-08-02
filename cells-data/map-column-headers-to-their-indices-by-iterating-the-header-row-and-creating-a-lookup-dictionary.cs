// Title: C# – Build a case‑insensitive header‑to‑column index dictionary with Aspose.Cells
// Description: Learn how to use Aspose.Cells for .NET to map column headers from the first worksheet row to zero‑based column indexes. The sample creates a Workbook, fills a header row, uses Cells.MaxDataColumn, iterates the row, skips empty cells, stores the mapping in a case‑insensitive Dictionary<string,int>, shows a lookup for "Price", and saves the file.
// Keywords: Aspose.Cells header mapping C# | column index dictionary .NET | iterate header row Aspose.Cells | MaxDataColumn property | case insensitive dictionary Excel | C# Excel header to index | Aspose.Cells example USA | Aspose.Cells tutorial UK | Aspose.Cells guide India
// Common Searches: Aspose.Cells map header to column index | C# get column number by header name | How to create header lookup dictionary in Aspose.Cells | Skip empty header cells Aspose.Cells | Retrieve column index for "Price" using Aspose.Cells
// Developer Intent: Generate a fast, case‑insensitive lookup that converts each non‑empty header in the first row of an Excel worksheet into its zero‑based column index using Aspose.Cells for .NET.
// Use Cases: Read data rows after locating the "Price" column to compute totals without hard‑coding column numbers. | Import CSV or user‑supplied Excel files where column order varies, by dynamically building a header‑to‑index map. | Write values back to specific columns (e.g., update inventory) using the dictionary instead of fixed indexes.
// AI Prompts: Write C# code that creates a case‑insensitive Dictionary<string,int> of header names to column indexes with Aspose.Cells, handling empty cells and using Cells.MaxDataColumn. | Show how to retrieve a column index for a given header and then read all cells under that column using Aspose.Cells in .NET. | Explain how to extend the header mapping to support duplicate header names by storing a List<int> for each header.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsHeaderLookup
{
    // Learn how to use Aspose.Cells for .NET to map column headers from the first worksheet row to zero‑based column indexes. The sample creates a Workbook, fills a header row, uses Cells.MaxDataColumn, iterates the row, skips empty cells, stores the mapping in a case‑insensitive Dictionary<string,int>, shows a lookup for "Price", and saves the file.
    public class HeaderMapper
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate a sample header row (row index 0)
                cells[0, 0].PutValue("ID");
                cells[0, 1].PutValue("Name");
                cells[0, 2].PutValue("Price");
                cells[0, 3].PutValue("Quantity");

                // Add some sample data rows (optional, just for completeness)
                cells[1, 0].PutValue(1);
                cells[1, 1].PutValue("Apple");
                cells[1, 2].PutValue(0.5);
                cells[1, 3].PutValue(100);

                cells[2, 0].PutValue(2);
                cells[2, 1].PutValue("Banana");
                cells[2, 2].PutValue(0.3);
                cells[2, 3].PutValue(150);

                // Create a dictionary to map header names to their column indices
                Dictionary<string, int> headerToIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                // Determine the last column that contains data in the header row
                int lastColumn = cells.MaxDataColumn; // correct Aspose.Cells property

                // Iterate through each column in the header row
                for (int col = 0; col <= lastColumn; col++)
                {
                    string header = cells[0, col].StringValue?.Trim();

                    // Skip empty headers
                    if (string.IsNullOrEmpty(header))
                        continue;

                    // Add or update the mapping
                    headerToIndex[header] = col;
                }

                // Demonstrate the lookup dictionary
                Console.WriteLine("Header to Column Index Mapping:");
                foreach (var kvp in headerToIndex)
                {
                    Console.WriteLine($"Header \"{kvp.Key}\" => Column Index {kvp.Value}");
                }

                // Example: retrieve the column index for a specific header
                if (headerToIndex.TryGetValue("Price", out int priceColumn))
                {
                    Console.WriteLine($"\nThe \"Price\" column is at index {priceColumn}.");
                }

                // Save the workbook (optional, just to illustrate lifecycle usage)
                workbook.Save("HeaderLookupDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for testing
    class Program
    {
        static void Main()
        {
            HeaderMapper.Run();
        }
    }
}
