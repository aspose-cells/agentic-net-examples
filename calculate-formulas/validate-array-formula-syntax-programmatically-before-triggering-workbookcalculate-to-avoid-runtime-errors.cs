// Title: Validate Array Formula Syntax with Aspose.Cells .NET Before Workbook.Calculate
// Description: C# sample that creates a workbook, inserts data, defines a legacy CSE array formula, programmatically checks each array formula using Sheet.CalculateArrayFormula, and runs Workbook.CalculateFormula only when all formulas are syntactically valid, avoiding runtime errors.
// Keywords: Aspose.Cells | C# | array formula validation | CalculateArrayFormula | Workbook.Calculate | IsArrayFormula | prevent runtime errors | Excel formula syntax check | programmatic formula testing | Aspose.Cells example
// Common Searches: Aspose.Cells validate array formula before calculation | C# CalculateArrayFormula example | how to check array formula syntax with Aspose.Cells | prevent invalid array formula errors in .NET | Aspose.Cells workbook.Calculate guard
// Developer Intent: Programmatically verify the syntax of every array formula in a workbook before invoking full calculation to eliminate runtime exceptions.
// Use Cases: Iterate through all cells, detect IsArrayFormula, and call Sheet.CalculateArrayFormula to confirm each formula parses correctly. | Abort Workbook.CalculateFormula when any array formula fails validation, then log the offending formulas. | Log validation results, display calculated array values after successful validation, and optionally save the workbook.
// AI Prompts: Write C# code that scans a worksheet for array formulas and uses Aspose.Cells CalculateArrayFormula to validate their syntax before calling Workbook.CalculateFormula. | Show how to catch exceptions from CalculateArrayFormula to identify invalid array formulas and stop workbook calculation in Aspose.Cells. | Provide an example that logs validation outcomes for each array formula and proceeds with full workbook calculation only when all are valid.

using System;
using Aspose.Cells;

namespace AsposeCellsArrayFormulaValidation
{
    // C# sample that creates a workbook, inserts data, defines a legacy CSE array formula, programmatically checks each array formula using Sheet.CalculateArrayFormula, and runs Workbook.CalculateFormula only when all formulas are syntactically valid, avoiding runtime errors.
    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a new workbook and get the first worksheet
            // -------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // -------------------------------------------------
            // 2. Populate some sample data that will be used by the array formulas
            // -------------------------------------------------
            cells["A1"].PutValue(1);
            cells["A2"].PutValue(2);
            cells["A3"].PutValue(3);
            cells["B1"].PutValue(10);
            cells["B2"].PutValue(20);
            cells["B3"].PutValue(30);

            // -------------------------------------------------
            // 3. Set an array formula (legacy CSE style) in cell C1
            //    This formula multiplies A1:A3 by B1:B3 element‑wise.
            // -------------------------------------------------
            Cell arrayCell = cells["C1"];
            // The result will occupy 3 rows × 1 column
            arrayCell.SetArrayFormula("=A1:A3*B1:B3", 3, 1);

            // -------------------------------------------------
            // 4. Validate all array formulas before any full workbook calculation
            // -------------------------------------------------
            bool allValid = true;
            foreach (Cell cell in cells)
            {
                if (cell.IsArrayFormula)
                {
                    try
                    {
                        // Attempt to calculate the array formula in isolation.
                        // The method returns a 2‑D object array; we only need to know that it succeeds.
                        object[][] dummyResult = sheet.CalculateArrayFormula(cell.Formula, new CalculationOptions());

                        // If we reach here, the formula syntax is valid.
                        Console.WriteLine($"Array formula in {cell.Name} is valid.");
                    }
                    catch (Exception ex)
                    {
                        // Syntax (or other) error detected.
                        allValid = false;
                        Console.WriteLine($"Invalid array formula in {cell.Name}: {ex.Message}");
                    }
                }
            }

            // -------------------------------------------------
            // 5. Only calculate the workbook if every array formula passed validation
            // -------------------------------------------------
            if (allValid)
            {
                workbook.CalculateFormula(); // Full workbook calculation
                Console.WriteLine("Workbook calculated successfully.");
            }
            else
            {
                Console.WriteLine("Workbook calculation aborted due to invalid array formulas.");
            }

            // -------------------------------------------------
            // 6. Output the results of the array formula (if calculated)
            // -------------------------------------------------
            if (allValid)
            {
                for (int i = 0; i < 3; i++)
                {
                    Console.WriteLine($"C{i + 1} = {cells[i, 2].Value}");
                }
            }

            // -------------------------------------------------
            // 7. Save the workbook (optional)
            // -------------------------------------------------
            workbook.Save("ArrayFormulaValidationResult.xlsx");
        }
    }
}
