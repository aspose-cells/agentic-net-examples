// Title: Ignore empty‑cell reference errors in Aspose.Cells formula calculation (C#)
// Description: Demonstrates how to suppress errors caused by formulas that reference empty cells in Aspose.Cells for .NET. The example creates a workbook, sets an empty value in A1, adds a formula in B1, configures CalculationOptions.IgnoreError, disables the EmptyCellRef check via ErrorCheckOptionCollection for a specific range, runs workbook.CalculateFormula, and saves the result.
// Keywords: Aspose.Cells | C# | CalculationOptions.IgnoreError | ErrorCheckOptionCollection | EmptyCellRef | ignore empty cell reference | suppress formula errors | formula calculation options | disable error check Aspose.Cells | Excel automation .NET
// Common Searches: Aspose.Cells ignore empty cell reference errors | How to disable EmptyCellRef error check in Aspose.Cells | CalculationOptions.IgnoreError example C# | Suppress formula errors in Aspose.Cells workbook | Apply error‑check options to a range Aspose.Cells
// Developer Intent: Prevent formula evaluation from throwing errors when referenced cells are blank.
// Use Cases: Run financial or statistical models where missing inputs should not halt calculation. | Generate Excel templates that tolerate empty cells without displaying error warnings. | Automate report creation that skips empty‑cell reference errors for defined sheet areas.
// AI Prompts: Show C# code to configure Aspose.Cells CalculationOptions to ignore all formula errors and turn off EmptyCellRef checks for a specific range. | Explain how CalculationOptions.IgnoreError and ErrorCheckOption.SetErrorCheck affect formula evaluation in Aspose.Cells. | Provide a step‑by‑step guide to suppress empty‑cell reference errors when calculating formulas with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Demonstrates how to suppress errors caused by formulas that reference empty cells in Aspose.Cells for .NET. The example creates a workbook, sets an empty value in A1, adds a formula in B1, configures CalculationOptions.IgnoreError, disables the EmptyCellRef check via ErrorCheckOptionCollection for a specific range, runs workbook.CalculateFormula, and saves the result.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Set up a scenario where a formula references an empty cell
        cells["A1"].PutValue("");               // Empty cell
        cells["B1"].Formula = "=A1*2";           // Formula that references the empty cell

        // Configure calculation options to ignore all errors during formula evaluation
        CalculationOptions calcOptions = new CalculationOptions
        {
            IgnoreError = true // Suppress errors such as those from empty‑cell references
        };

        // Additionally, disable the specific error check for empty‑cell references
        ErrorCheckOptionCollection errorCheckOptions = worksheet.ErrorCheckOptions;
        int optionIndex = errorCheckOptions.Add();               // Add a new error‑check option
        ErrorCheckOption errorCheckOption = errorCheckOptions[optionIndex];
        errorCheckOption.SetErrorCheck(ErrorCheckType.EmptyCellRef, false); // Do not flag empty‑cell reference errors
        errorCheckOption.AddRange(CellArea.CreateCellArea("A1", "Z100"));   // Apply to the desired range

        // Perform calculation with the configured options
        workbook.CalculateFormula(calcOptions);

        // Display the result of the formula that referenced the empty cell
        Console.WriteLine("Result in B1: " + cells["B1"].StringValue);

        // Save the workbook
        workbook.Save("IgnoreEmptyCellReference.xlsx");
    }
}
