// Title: Insert a SORT dynamic array formula into B2 and validate the spilled range size with Aspose.Cells for .NET
// AI Prompts: Place a =SORT(A2:A6) formula in cell B2 using Aspose.Cells, trigger workbook calculation, then programmatically count the non‑empty cells in column B to ensure the spilled range length matches the original array size. | Write C# code that adds a SORT dynamic array formula to B2, evaluates the workbook, iterates down the spilled column, and asserts that the number of returned rows equals the source data count.
// Common Searches: asp.net insert SORT dynamic array formula with Aspose.Cells | how to verify spilled range size after dynamic array calculation in C# | Aspose.Cells read spilled cells from a SORT formula | C# count rows returned by a dynamic array formula using Aspose.Cells | validate that SORT(A2:A6) spill matches source range length in .NET
// Tags: insert SORT formula Aspose.Cells | validate spilled range C# | dynamic array calculation Aspose.Cells | read spilled cells after formula | compare spilled rows to source data

using System;
using Aspose.Cells;

// The example creates a workbook, fills A2:A6 with numbers, inserts the =SORT(A2:A6) dynamic array formula into B2, calculates the workbook, scans column B from B2 downward to count non‑empty cells, compares this count to the original data length, reports whether the spilled range expands correctly, and saves the file as DynamicArraySort.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate source data in A2:A6
            int[] numbers = { 5, 2, 9, 1, 7 };
            for (int i = 0; i < numbers.Length; i++)
            {
                // Cells[row, column] - zero based indexing
                sheet.Cells[i + 1, 0].PutValue(numbers[i]); // A2:A6
            }

            // Insert SORT dynamic array formula in B2
            Cell formulaCell = sheet.Cells[1, 1]; // B2
            formulaCell.Formula = "=SORT(A2:A6)";

            // Calculate the workbook to evaluate the formula
            workbook.CalculateFormula();

            // Determine the spilled range size by scanning column B starting at B2
            int startRow = 1; // zero‑based index for row 2
            int startCol = 1; // column B
            int actualRows = 0;
            while (true)
            {
                Cell c = sheet.Cells[startRow + actualRows, startCol];
                if (c.Value == null || string.IsNullOrEmpty(c.StringValue))
                    break;
                actualRows++;
            }

            // Verify that the spilled range size matches source data count
            int expectedRows = numbers.Length;

            Console.WriteLine($"Expected rows: {expectedRows}, Actual rows: {actualRows}");

            if (actualRows == expectedRows)
                Console.WriteLine("Spilled range expands correctly.");
            else
                Console.WriteLine("Spilled range size mismatch.");

            // Save the workbook (optional)
            workbook.Save("DynamicArraySort.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
