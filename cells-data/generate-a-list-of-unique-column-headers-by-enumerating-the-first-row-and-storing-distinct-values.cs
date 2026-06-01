using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsUniqueHeadersDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate the first row with sample headers (some duplicates)
            cells["A1"].PutValue("Name");
            cells["B1"].PutValue("Age");
            cells["C1"].PutValue("Name");   // duplicate
            cells["D1"].PutValue("Email");
            cells["E1"].PutValue("Age");    // duplicate
            cells["F1"].PutValue("Country");

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

            // Convert the set to a list (preserves no particular order)
            List<string> uniqueHeaders = new List<string>(uniqueHeadersSet);

            // Output the unique headers
            Console.WriteLine("Unique column headers:");
            foreach (string header in uniqueHeaders)
            {
                Console.WriteLine(header);
            }

            // Optionally, save the workbook to verify the data (not required for header extraction)
            workbook.Save("UniqueHeadersDemo.xlsx");
        }
    }
}