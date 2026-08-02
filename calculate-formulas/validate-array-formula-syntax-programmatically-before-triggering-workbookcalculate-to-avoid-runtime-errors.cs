// Title: C# – Validate Array Formula Syntax in Aspose.Cells Before Workbook.Calculate
// Description: Demonstrates how to programmatically scan a worksheet, test each legacy array formula with Worksheet.CalculateArrayFormula, catch syntax errors, log the offending cell, clear invalid formulas, and then safely call workbook.CalculateFormula. Prevents runtime exceptions caused by malformed array formulas.
// Keywords: Aspose.Cells array formula validation | C# calculate array formula | Worksheet.CalculateArrayFormula | prevent runtime errors Aspose.Cells | validate Excel formulas .NET | clear invalid array formulas | Excel automation C# | Aspose.Cells workbook.Calculate safety
// Common Searches: validate array formulas Aspose.Cells C# | catch invalid formula syntax before workbook.Calculate | remove bad array formulas programmatically | Aspose.Cells calculate formula error handling | how to test Excel formula syntax with Aspose.Cells
// Developer Intent: Ensure all array formulas are syntactically correct before triggering workbook.Calculate to avoid calculation failures.
// Use Cases: Iterate over a worksheet’s used range, validate each array formula, and clear those that throw exceptions. | Log cell addresses and error messages for malformed array formulas to aid debugging of generated reports. | Integrate automatic formula validation into batch Excel generation pipelines so only valid formulas are retained.
// AI Prompts: Create a C# method that scans a worksheet, validates each array formula using Aspose.Cells, logs invalid entries, and removes them before calculation. | Show how to catch exceptions from Worksheet.CalculateArrayFormula for malformed formulas and safely continue processing. | Write an example that demonstrates safe workbook calculation after programmatically clearing invalid array formulas.

using System;
using Aspose.Cells;

// Demonstrates how to programmatically scan a worksheet, test each legacy array formula with Worksheet.CalculateArrayFormula, catch syntax errors, log the offending cell, clear invalid formulas, and then safely call workbook.CalculateFormula. Prevents runtime exceptions caused by malformed array formulas.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample data for demonstration
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].PutValue(30);

            // Set a valid array formula
            cells["B1"].SetArrayFormula("=SUM(A1:A3)", 1, 1);

            // Set an invalid array formula (syntax error)
            cells["C1"].SetArrayFormula("=SUM(A1:A3", 1, 1); // missing closing parenthesis

            // Validate all array formulas before any calculation
            ValidateArrayFormulas(worksheet);

            // Calculate all formulas after validation (lifecycle rule: calculate)
            workbook.CalculateFormula();

            // Save the workbook (lifecycle rule: save)
            workbook.Save("ValidatedResult.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }

    // Checks each cell that contains an array formula.
    // If the formula cannot be parsed/calculated, it is cleared to prevent runtime errors.
    static void ValidateArrayFormulas(Worksheet worksheet)
    {
        Cells cells = worksheet.Cells;

        // Determine the used range to limit iteration
        int maxRow = cells.MaxDataRow;
        int maxCol = cells.MaxDataColumn;

        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = cells[row, col];

                // Process only legacy array formulas (IsArrayFormula == true)
                if (cell.IsArrayFormula)
                {
                    try
                    {
                        // Attempt to calculate the array formula.
                        // If the syntax is invalid, an exception will be thrown.
                        worksheet.CalculateArrayFormula(cell.Formula, new CalculationOptions());
                    }
                    catch (Exception ex)
                    {
                        // Log the problematic cell and error details
                        Console.WriteLine($"Invalid array formula at {cell.Name}: {cell.Formula}");
                        Console.WriteLine($"Error: {ex.Message}");

                        // Remove the bad formula to avoid calculation failures
                        cell.Formula = string.Empty;
                    }
                }
            }
        }
    }
}
