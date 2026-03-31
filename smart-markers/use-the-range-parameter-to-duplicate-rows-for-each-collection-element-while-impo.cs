using System;
using System.Collections.Generic;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX file (uses the provided load rule)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Collection whose elements will be written to duplicated rows
        // Each object[] represents the values for a single row
        List<object[]> collection = new List<object[]>
        {
            new object[] { "John", 30, "New York" },
            new object[] { "Jane", 25, "London" },
            new object[] { "Bob", 35, "Paris" }
        };

        // Index of the template row that contains the desired formatting
        // (0‑based; assume row 1 is the template)
        int templateRowIndex = 1;

        // Current row pointer – starts at the template row
        int currentRow = templateRowIndex;

        foreach (object[] rowData in collection)
        {
            // Insert a new row directly below the current row (uses InsertRows rule)
            cells.InsertRows(currentRow + 1, 1, true);

            // Copy formatting from the template row to the newly inserted row (uses CopyRows rule)
            cells.CopyRows(cells, templateRowIndex, currentRow + 1, 1);

            // Import the data horizontally into the new row (uses ImportObjectArray rule)
            cells.ImportObjectArray(rowData, currentRow + 1, 0, false);

            // Advance the pointer to the row just filled
            currentRow++;
        }

        // Save the modified workbook (uses the provided save rule)
        workbook.Save("output.xlsx");
    }
}