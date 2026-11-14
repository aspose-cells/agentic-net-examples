// Title: How to evaluate a SUM formula and read the resulting value from cell E3 using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get the value of a cell after workbook.CalculateFormula | How to read the result of =SUM(A1:D1) placed in E3 with Aspose.Cells | Retrieve numeric result of a formula cell in Aspose.Cells after recalculation | Aspose.Cells example to calculate formulas and access cell E3 value
// Tags: SUM formula evaluation Aspose.Cells C# | access calculated cell value Aspose.Cells | workbook.CalculateFormula usage .NET | read formula result from worksheet cell | retrieve cell value after formula recalculation

using Aspose.Cells;
using System;

// The example creates a workbook, assigns numbers to cells A1‑D1, sets a SUM formula in E3, runs workbook.CalculateFormula to evaluate all formulas, then reads the computed value of E3 and prints it to the console.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some cells that will be used in a formula
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["B1"].PutValue(20);
        sheet.Cells["C1"].PutValue(30);
        sheet.Cells["D1"].PutValue(40);

        // Set a formula in E3 (e.g., sum of A1:D1)
        sheet.Cells["E3"].Formula = "=SUM(A1:D1)";

        // Recalculate all formulas in the workbook
        workbook.CalculateFormula();

        // Retrieve the final calculated value of cell E3
        object e3Value = sheet.Cells["E3"].Value;

        // Display the value
        Console.WriteLine("Calculated value of E3: " + e3Value);
    }
}
