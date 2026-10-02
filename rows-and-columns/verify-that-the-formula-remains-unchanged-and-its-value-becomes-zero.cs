// Title: Check that an Aspose.Cells formula string stays the same while its calculated value becomes zero in C#
// AI Prompts: Generate C# code using Aspose.Cells that sets a formula in a cell, forces workbook recalculation, and asserts that the formula text has not changed after calculation. | Write a verification routine in C# with Aspose.Cells that confirms a cell's evaluated result is zero while preserving the original formula expression.
// Common Searches: how to ensure a formula text is unchanged after Aspose.Cells workbook.CalculateFormula in C# | Aspose.Cells C# verify that cell B1 formula result is zero when referenced cells contain zeros | unit test for Aspose.Cells formula integrity and zero result in .NET | C# Aspose.Cells check that formula string remains 'A1+A2' after recalculation
// Tags: Aspose.Cells formula integrity after CalculateFormula | C# check zero result of Excel formula with Aspose.Cells | Aspose.Cells workbook recalculation unit test | Validate Excel formula string preservation using Aspose.Cells | Aspose.Cells evaluate dependent cells returning zero

using System;
using Aspose.Cells;

// The example creates a workbook, places zeros in A1 and A2, assigns the formula "A1+A2" to B1, recalculates the workbook, and then verifies that the formula text remains "A1+A2" and that the evaluated value in B1 is zero.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Set up cells that will be used in the formula
        sheet.Cells["A1"].PutValue(0);
        sheet.Cells["A2"].PutValue(0);

        // Define a formula in cell B1
        Cell formulaCell = sheet.Cells["B1"];
        string originalFormula = "A1+A2";
        formulaCell.Formula = originalFormula;

        // Recalculate the workbook to evaluate the formula
        workbook.CalculateFormula();

        // Verify that the formula text remains unchanged
        if (formulaCell.Formula != originalFormula)
        {
            Console.WriteLine("Formula has changed!");
        }
        else
        {
            Console.WriteLine("Formula unchanged: " + formulaCell.Formula);
        }

        // Verify that the evaluated value of the formula is zero
        if (formulaCell.Value is double numericValue && Math.Abs(numericValue) < 1e-10)
        {
            Console.WriteLine("Formula value is zero as expected.");
        }
        else
        {
            Console.WriteLine("Formula value is not zero. Value: " + formulaCell.Value);
        }

        // Optionally save the workbook (not required for verification)
        // workbook.Save("VerificationResult.xlsx");
    }
}
