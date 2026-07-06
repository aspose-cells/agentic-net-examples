using System;
using Aspose.Cells;

namespace AsposeCellsMinifsValidation
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Set workbook to allow newer functions (Excel 2016+)
            // Disabling compatibility check ensures functions like MINIFS are retained.
            workbook.Settings.CheckCompatibility = false;

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            // Column A: Values to evaluate
            // Column B: Criteria values
            double[] values = { 5, 12, 8, 20, 15 };
            double[] criteria = { 7, 12, 9, 20, 5 };

            for (int i = 0; i < values.Length; i++)
            {
                cells[i + 1, 0].PutValue(values[i]);   // A2:A6
                cells[i + 1, 1].PutValue(criteria[i]); // B2:B6
            }

            // Set MINIFS formula:
            // =MINIFS(A2:A6, B2:B6, ">=10")
            // This should return the minimum value in A2:A6 where corresponding B2:B6 >= 10.
            cells["C2"].Formula = "=MINIFS(A2:A6, B2:B6, \">=10\")";

            // Calculate formulas (lifecycle rule)
            workbook.CalculateFormula();

            // Retrieve and display the result
            double result = cells["C2"].DoubleValue;
            Console.WriteLine("MINIFS result (expected 12): " + result);

            // Simple validation
            if (Math.Abs(result - 12) < 0.0001)
                Console.WriteLine("Validation passed.");
            else
                Console.WriteLine("Validation failed.");

            // Save the workbook (lifecycle rule)
            workbook.Save("MinifsValidation.xlsx");
        }
    }
}