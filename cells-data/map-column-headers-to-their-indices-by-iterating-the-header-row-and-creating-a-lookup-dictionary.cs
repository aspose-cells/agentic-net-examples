using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace HeaderMappingExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample header row (row index 0)
            cells["A1"].PutValue("Name");
            cells["B1"].PutValue("Age");
            cells["C1"].PutValue("Country");
            cells["D1"].PutValue("Salary");

            // Populate some data rows (optional, just for illustration)
            cells["A2"].PutValue("John");
            cells["B2"].PutValue(30);
            cells["C2"].PutValue("USA");
            cells["D2"].PutValue(50000);

            cells["A3"].PutValue("Alice");
            cells["B3"].PutValue(28);
            cells["C3"].PutValue("UK");
            cells["D3"].PutValue(48000);

            // Create a lookup dictionary that maps header names to their column indices
            Dictionary<string, int> headerToIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            // Determine the last column that contains data
            int lastColumn = cells.MaxDataColumn;

            // Iterate over the header row (row 0) and fill the dictionary
            for (int col = 0; col <= lastColumn; col++)
            {
                string header = cells[0, col].StringValue;
                if (!string.IsNullOrEmpty(header))
                {
                    // Store the column index for the header name
                    headerToIndex[header] = col;
                }
            }

            // Example usage: retrieve the column index for a specific header
            if (headerToIndex.TryGetValue("Age", out int ageColumnIndex))
            {
                Console.WriteLine($"Column index for 'Age' is {ageColumnIndex}");
                // Access a cell using the retrieved column index
                Console.WriteLine($"Value in row 2, 'Age' column: {cells[1, ageColumnIndex].IntValue}");
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save("HeaderMappingDemo.xlsx");
        }
    }
}