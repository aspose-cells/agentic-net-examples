// Title: Ignore formula errors during calculation with Aspose.Cells CalculationOptions in C#
// AI Prompts: Generate C# code that creates a workbook, sets a formula that would cause an error, enables CalculationOptions.IgnoreError, and runs workbook.CalculateFormula without throwing exceptions. | Show how to configure Aspose.Cells to skip cells that return errors (e.g., division by zero) when evaluating all formulas in a .NET workbook. | Provide a complete example that calculates formulas with error suppression and then saves the workbook to an .xlsx file using Aspose.Cells.
// Common Searches: aspocells calculationoptions ignoreerror true example c# | how to prevent #DIV/0! exception when calling workbook.CalculateFormula in Aspose.Cells | skip formula errors during workbook calculation Aspose.Cells .NET | ignore errors while evaluating formulas in an Excel file using Aspose.Cells C# | calculate all formulas without raising exceptions Aspose.Cells C#
// Tags: Aspose.Cells calculationoptions ignoreerror | C# error‑tolerant formula evaluation Aspose.Cells | skip division by zero errors Aspose.Cells | calculate workbook formulas without exceptions .NET | error suppression during Excel formula calculation Aspose.Cells | ignore formula errors when saving workbook C#

using System;
using Aspose.Cells;

// Creates a workbook, writes a division‑by‑zero formula, sets CalculationOptions.IgnoreError to true, calculates all formulas without raising an exception, outputs the cell value, and saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Set a formula that would normally cause an error (division by zero)
        sheet.Cells["A1"].Formula = "=1/0";

        // Create calculation options and enable ignoring errors
        CalculationOptions calcOptions = new CalculationOptions
        {
            IgnoreError = true
        };

        // Calculate all formulas using the specified options
        workbook.CalculateFormula(calcOptions);

        // Display the result; no exception is thrown even though the formula is erroneous
        Console.WriteLine("A1 value after calculation: " + sheet.Cells["A1"].StringValue);

        // Save the workbook (optional)
        workbook.Save("IgnoreErrorDemo.xlsx");
    }
}
