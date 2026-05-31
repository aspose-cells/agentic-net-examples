using System;
using Aspose.Cells;

namespace AsposeCellsIFNADemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Example 1: A1 has a valid value, IFNA should return that value
                cells["A1"].PutValue(42);
                cells["B1"].Formula = "=IFNA(A1, \"fallback\")";

                // Calculate all formulas in the workbook
                workbook.CalculateFormula();

                // Display the result of IFNA when A1 is not an error
                Console.WriteLine("B1 (A1 = 42) => " + cells["B1"].StringValue); // Expected: 42

                // Example 2: Clear A1 so the formula evaluates to an error (empty cell is treated as #N/A)
                cells["A1"].PutValue(string.Empty); // Clear contents

                // Recalculate formulas
                workbook.CalculateFormula();

                // Display the result of IFNA when A1 is empty (fallback should be used)
                Console.WriteLine("B1 (A1 empty) => " + cells["B1"].StringValue); // Expected: fallback
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}