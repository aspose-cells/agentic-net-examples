using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle create rule)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Formula that references an empty cell (B1 is empty)
        cells["A1"].Formula = "=B1+10";

        // Configure error checking to ignore errors caused by empty‑cell references
        // The ErrorCheckOptionCollection holds per‑range error‑check settings.
        ErrorCheckOptionCollection errorCheckOptions = worksheet.ErrorCheckOptions;
        int optionIndex = errorCheckOptions.Add();                     // add a new option
        ErrorCheckOption errorCheckOption = errorCheckOptions[optionIndex];

        // Disable the EmptyCellRef check – this tells Aspose.Cells to ignore such errors.
        errorCheckOption.SetErrorCheck(ErrorCheckType.EmptyCellRef, false);

        // Apply the option to the cell(s) that contain the formula.
        errorCheckOption.AddRange(CellArea.CreateCellArea("A1", "A1"));

        // Optionally, also set the global calculation option to ignore all errors.
        CalculationOptions calcOptions = new CalculationOptions
        {
            IgnoreError = true
        };

        // Calculate all formulas with the configured options.
        workbook.CalculateFormula(calcOptions);

        // Save the workbook (lifecycle save rule)
        workbook.Save("IgnoreEmptyCellReference.xlsx");
    }
}