// Title: Get a list of unique column headers from the first row of an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that reads the first row of a worksheet and returns a collection of distinct header strings, ignoring case. | Show how to use a HashSet<string> to eliminate duplicate column names while iterating over cells in Aspose.Cells. | Provide an example that prints the unique headers and saves the workbook after processing.
// Common Searches: aspnet c# how to retrieve distinct column names from first row using Aspose.Cells | remove duplicate header values in Excel file with Aspose.Cells .NET API | enumerate cells in first row and get unique headers Aspose.Cells example | max data column first row Aspose.Cells get header list
// Tags: unique header extraction Aspose.Cells | first row column enumeration .NET | deduplication of Excel headers using .NET collection | retrieve last populated column Aspose.Cells | persist workbook after header extraction

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The sample creates a workbook, adds sample headers (including duplicates) to the first row, determines the last populated column, iterates over the first row collecting non‑empty header values into a case‑insensitive HashSet, converts the set to a list preserving insertion order, prints each unique header, and finally saves the workbook as UniqueHeadersDemo.xlsx.
    class UniqueColumnHeadersDemo
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample data: first row contains headers (some duplicates)
            cells["A1"].PutValue("Name");
            cells["B1"].PutValue("Age");
            cells["C1"].PutValue("Name");   // duplicate header
            cells["D1"].PutValue("Email");
            cells["E1"].PutValue("Age");    // duplicate header

            // Determine the last column that contains data in the first row
            int lastColumn = cells.MaxDataColumn;

            // Use a HashSet to collect distinct header values
            HashSet<string> uniqueHeadersSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int col = 0; col <= lastColumn; col++)
            {
                string header = cells[0, col].StringValue?.Trim();
                if (!string.IsNullOrEmpty(header))
                {
                    uniqueHeadersSet.Add(header);
                }
            }

            // Convert the set to a list (preserves insertion order in .NET 6+)
            List<string> uniqueHeaders = new List<string>(uniqueHeadersSet);

            // Display the unique headers
            Console.WriteLine("Unique column headers:");
            foreach (string header in uniqueHeaders)
            {
                Console.WriteLine(header);
            }

            // Save the workbook (optional, just to demonstrate lifecycle usage)
            workbook.Save("UniqueHeadersDemo.xlsx");
        }
    }
}
