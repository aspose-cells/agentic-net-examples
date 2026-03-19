using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet's cells collection
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // ------------------------------------------------------------
            // 1. Create a range from C2 to E5 using the CreateRange(string, string) method
            // ------------------------------------------------------------
            AsposeRange myRange = cells.CreateRange("C2", "E5");

            // Give the range a name (optional but useful for formulas)
            myRange.Name = "DataBlock";

            // Fill the range with sample data
            for (int row = 0; row < myRange.RowCount; row++)
            {
                for (int col = 0; col < myRange.ColumnCount; col++)
                {
                    // Example value: "R{row}C{col}"
                    myRange[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // ------------------------------------------------------------
            // 2. Add the created range to the Cells collection using AddRange(Range)
            // ------------------------------------------------------------
            cells.AddRange(myRange);

            // ------------------------------------------------------------
            // 3. Demonstrate that the range expands when rows/columns are inserted
            // ------------------------------------------------------------
            // Insert a new row above the range (row index 1 corresponds to Excel row 2)
            cells.InsertRow(1); // This shifts existing rows down, expanding the range

            // Insert a new column to the left of the range (column index 2 corresponds to column C)
            cells.InsertColumn(2); // This shifts existing columns right, expanding the range

            // After insertion, the range object automatically reflects the new size
            Console.WriteLine($"Range address after insertions: {myRange.Address}");
            Console.WriteLine($"Rows: {myRange.RowCount}, Columns: {myRange.ColumnCount}");

            // ------------------------------------------------------------
            // 4. Use the range in a simple formula (sum of the first column of the range)
            // ------------------------------------------------------------
            // Place formula in cell B1 (outside the range)
            cells["B1"].Formula = $"=SUM({myRange.Name})";

            // Calculate formulas
            workbook.CalculateFormula();

            // Output the calculated sum
            Console.WriteLine($"Sum of first column in '{myRange.Name}': {cells["B1"].DoubleValue}");

            // ------------------------------------------------------------
            // 5. Save the workbook as XLSX
            // ------------------------------------------------------------
            workbook.Save("AdvancedRangeDemo.xlsx");
        }
    }
}