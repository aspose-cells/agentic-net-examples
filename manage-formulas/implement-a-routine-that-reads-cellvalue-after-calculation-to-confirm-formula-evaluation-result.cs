// Title: How to read a cell's calculated value after using Workbook.CalculateFormula with Aspose.Cells in C#
// Common Searches: Aspose.Cells C# get value of a formula cell after workbook.CalculateFormula | read calculated result from B1 after setting formula =A1+A2 using Aspose.Cells | C# example for retrieving Excel formula output with Aspose.Cells API | how to access cell.Value after recalculating formulas in Aspose.Cells workbook | Aspose.Cells calculate all formulas then read result programmatically C#
// Tags: Aspose.Cells calculate formula C# | retrieve cell value after Workbook.CalculateFormula | read formula result from worksheet Aspose.Cells | C# Excel formula evaluation example | Workbook.CalculateFormula usage Aspose.Cells

using Aspose.Cells;
using System;

// The sample creates a Workbook, writes numbers to A1 and A2, assigns the formula "=A1+A2" to B1, calls Workbook.CalculateFormula to evaluate all formulas, then reads the computed result via B1.Value and prints it.
class FormulaEvaluation
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate cells with sample data
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);

        // Set a formula in B1 that adds A1 and A2
        sheet.Cells["B1"].Formula = "=A1+A2";

        // Recalculate all formulas in the workbook
        workbook.CalculateFormula();

        // Retrieve the calculated value from B1
        object result = sheet.Cells["B1"].Value;

        // Display the result
        Console.WriteLine($"Calculated value in B1: {result}");
    }
}
