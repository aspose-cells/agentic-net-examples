// Title: Programmatically validate array formula syntax in an Excel workbook with Aspose.Cells for .NET before running Workbook.CalculateFormula
// AI Prompts: Generate C# code that iterates all worksheets in a workbook, uses Sheet.CalculateArrayFormula to test each array formula for syntax errors, collects invalid cell addresses, and only calls Workbook.CalculateFormula when no errors are found. | Create a reusable C# method that accepts a Workbook object, validates every array formula with custom CalculationOptions, and returns a list of cells containing invalid formulas. | Show how to configure CalculationOptions (e.g., EnableIterativeCalculation, PrecisionAsDisplayed) while validating array formulas and then perform full workbook calculation if validation passes.
// Common Searches: Aspose.Cells .NET how to check array formula syntax before workbook.CalculateFormula | C# validate Excel array formulas programmatically using Aspose.Cells | prevent runtime exception from invalid array formulas in Aspose.Cells workbook calculation | scan Excel workbook for array formulas and test them with Aspose.Cells API
// Tags: array formula syntax validation Aspose.Cells | calculate array formula with CalculationOptions .NET | iterate worksheets to detect invalid array formulas C# | conditional workbook calculation after formula check Aspose | exception handling for invalid array formulas Aspose.Cells

using System;
using Aspose.Cells;

// Loads a workbook, walks through each worksheet's used range, attempts to calculate each cell's array formula with Sheet.CalculateArrayFormula to catch syntax errors, aborts full workbook calculation if any invalid formulas are found, and saves the workbook.
class ValidateArrayFormulas
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        bool allFormulasValid = true;

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Determine the used range to limit the iteration
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Scan every cell within the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Process only cells that contain an array formula
                    if (cell.IsArrayFormula)
                    {
                        string formula = cell.Formula;

                        try
                        {
                            // Attempt to calculate the array formula.
                            // If the syntax is invalid, an exception will be thrown.
                            sheet.CalculateArrayFormula(formula, new CalculationOptions());
                        }
                        catch (Exception ex)
                        {
                            allFormulasValid = false;
                            Console.WriteLine($"Invalid array formula at {cell.Name}: {formula}");
                            Console.WriteLine($"Error: {ex.Message}");
                        }
                    }
                }
            }
        }

        // If all array formulas are syntactically correct, perform full workbook calculation
        if (allFormulasValid)
        {
            workbook.CalculateFormula();
            Console.WriteLine("Workbook calculated successfully.");
        }
        else
        {
            Console.WriteLine("Workbook contains invalid array formulas. Calculation aborted.");
        }

        // Save the workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
