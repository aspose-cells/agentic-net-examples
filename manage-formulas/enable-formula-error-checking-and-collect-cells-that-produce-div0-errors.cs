using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsErrorCheckDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // -------------------------------------------------
            // 1. Set up formulas that will generate #DIV/0! errors
            // -------------------------------------------------
            cells["A1"].PutValue(10);          // divisor
            cells["B1"].PutValue(0);           // divisor zero
            cells["C1"].Formula = "=A1/B1";    // #DIV/0! error
            cells["D1"].Formula = "=10/0";     // #DIV/0! error
            cells["E1"].Formula = "=SUM(A1:B1)"; // valid formula

            // -------------------------------------------------
            // 2. Enable error checking for evaluation errors (e.g., #DIV/0!)
            // -------------------------------------------------
            ErrorCheckOptionCollection errorCheckOptions = worksheet.ErrorCheckOptions;
            int optionIndex = errorCheckOptions.Add();
            ErrorCheckOption errorCheckOption = errorCheckOptions[optionIndex];

            // Enable checking for evaluation errors (green triangle will appear in Excel)
            errorCheckOption.SetErrorCheck(ErrorCheckType.EvaluationError, true);

            // Apply the check to the whole used range of the worksheet
            CellArea usedArea = CellArea.CreateCellArea(0, 0, cells.MaxRow, cells.MaxColumn);
            errorCheckOption.AddRange(usedArea);

            // -------------------------------------------------
            // 3. Calculate all formulas (do not ignore errors)
            // -------------------------------------------------
            workbook.CalculateFormula(new CalculationOptions() { IgnoreError = false });

            // -------------------------------------------------
            // 4. Collect addresses of cells that contain #DIV/0! errors
            // -------------------------------------------------
            List<string> divZeroCells = new List<string>();

            // Iterate through all cells that have data
            for (int row = 0; row <= cells.MaxRow; row++)
            {
                for (int col = 0; col <= cells.MaxColumn; col++)
                {
                    Cell cell = cells[row, col];
                    if (cell.IsErrorValue)
                    {
                        // Get the rich value to determine the specific error type
                        CellRichValue richValue = cell.GetRichValue();
                        if (richValue != null && richValue.ErrorValue == ErrorCellValueType.Calc)
                        {
                            // #DIV/0! is represented by the Calc error type
                            divZeroCells.Add(cell.Name);
                        }
                    }
                }
            }

            // Output the collected cell addresses
            Console.WriteLine("Cells containing #DIV/0! errors:");
            foreach (string address in divZeroCells)
            {
                Console.WriteLine(address);
            }

            // -------------------------------------------------
            // 5. Save the workbook (lifecycle rule)
            // -------------------------------------------------
            workbook.Save("ErrorCheckDivZeroDemo.xlsx");
        }
    }
}