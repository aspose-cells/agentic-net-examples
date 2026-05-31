using System;
using Aspose.Cells;

namespace AsposeCellsArrayFormulaDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (3 rows, 2 columns)
            cells["A1"].PutValue(1);
            cells["A2"].PutValue(2);
            cells["A3"].PutValue(3);
            cells["B1"].PutValue(4);
            cells["B2"].PutValue(5);
            cells["B3"].PutValue(6);

            // Define an array formula that aggregates (sums) all values in the range A1:B3
            string arrayFormula = "=SUM(A1:B3)";

            // Create calculation options (default options are sufficient for this example)
            CalculationOptions calcOptions = new CalculationOptions();

            // Calculate the array formula and obtain the result as a 2‑dimensional object array
            // Using the overload that allows specifying maximum dimensions (optional here)
            object[][] result = sheet.CalculateArrayFormula(arrayFormula, calcOptions, 1, 1);

            // The result array contains a single element with the sum of the range
            Console.WriteLine("Array formula result (SUM of A1:B3): " + result[0][0]);

            // Optionally, write the result back to a cell for verification
            cells["C1"].PutValue(result[0][0]);

            // Save the workbook to demonstrate that the data and result are persisted
            workbook.Save("ArrayFormulaResult.xlsx");
        }
    }
}