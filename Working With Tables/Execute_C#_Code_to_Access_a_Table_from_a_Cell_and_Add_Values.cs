using Aspose.Cells;
using System;

class Program
{
    static void Main()
    {
        // Load the existing XLSX file
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Create a range that starts at B2 and ends at D6 (5 rows x 3 columns)
        Aspose.Cells.Range range = cells.CreateRange("B2", "D6");

        // Fill the range using row and column offsets
        for (int rowOffset = 0; rowOffset < 5; rowOffset++)          // 5 rows
        {
            for (int colOffset = 0; colOffset < 3; colOffset++)      // 3 columns
            {
                // Example value: (rowIndex + 1) * (colIndex + 1)
                range[rowOffset, colOffset].PutValue((rowOffset + 1) * (colOffset + 1));
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}