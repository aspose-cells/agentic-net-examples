using System;
using Aspose.Cells;

namespace AsposeCellsFreezeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the Workbook() constructor rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate a 10x10 data matrix with sample values
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    // Example value: row index * 10 + column index
                    sheet.Cells[row, col].PutValue(row * 10 + col);
                }
            }

            // Freeze the top‑left 5 × 5 area
            // Parameters: row index, column index, number of frozen rows, number of frozen columns
            // Using FreezePanes(int, int, int, int) method as defined in the documentation
            sheet.FreezePanes(5, 5, 5, 5);

            // Save the workbook to a file (uses the Workbook.Save(string) rule)
            workbook.Save("FreezeTopLeft5x5.xlsx");
        }
    }
}