using System;
using Aspose.Cells;

namespace FormulaEvaluationExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Get the first worksheet (or specify the desired index)
            Worksheet sheet = workbook.Worksheets[0];

            // Prepare calculation options (optional, but ensures recursive calculation)
            CalculationOptions calcOptions = new CalculationOptions
            {
                Recursive = true   // Calculate dependent formulas in other worksheets if any
            };

            // Evaluate all formulas in the worksheet
            sheet.CalculateFormula(calcOptions, true);

            // Retrieve the calculated value of cell G10
            Cell targetCell = sheet.Cells["G10"];
            object g10Value = targetCell.Value;

            // Output the result
            Console.WriteLine($"Calculated value of G10: {g10Value}");

            // (Optional) Save the workbook if you want to persist the calculated results
            // workbook.Save("output.xlsx");
        }
    }
}