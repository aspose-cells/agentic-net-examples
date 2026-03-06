using System;
using Aspose.Cells;

namespace DecreaseCellCalculateTime
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // 1. Create LoadOptions and disable formula parsing on open.
            //    This skips the expensive parsing step when the workbook is loaded,
            //    which reduces the overall time before any calculation is performed.
            LoadOptions loadOptions = new LoadOptions
            {
                ParsingFormulaOnOpen = false
            };

            // 2. Load the workbook with the specified options.
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // 3. Access the cell whose formula we want to calculate.
            //    For demonstration we use cell A1.
            Worksheet sheet = workbook.Worksheets[0];
            Cell targetCell = sheet.Cells["A1"];

            // 4. Prepare CalculationOptions to speed up Cell.Calculate.
            //    - Set Recursive to false so only the target cell is evaluated,
            //      without traversing the whole dependency tree.
            //    - Reduce the recursion stack size (optional) to avoid overhead.
            CalculationOptions calcOptions = new CalculationOptions
            {
                Recursive = false,
                CalcStackSize = 50   // smaller than default 200, suitable for simple formulas
            };

            // 5. Perform the calculation on the single cell.
            targetCell.Calculate(calcOptions);

            // 6. Optionally, save the workbook after calculation.
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            // 7. Display the calculated value.
            Console.WriteLine($"Calculated value of A1: {targetCell.Value}");
        }
    }
}