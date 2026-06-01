using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook from disk
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet and its cells collection
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Identify the cell that contains the original formula (e.g., B2)
        Cell formulaCell = cells["B2"];

        // Show the original formula and its calculated value
        Console.WriteLine("Original formula: " + formulaCell.Formula);
        Console.WriteLine("Original value: " + formulaCell.Value);

        // Modify the formula to reference a different cell.
        // Example: change from "=A1*2" to "=C1*2"
        string newFormula = "=C1*2";
        formulaCell.Formula = newFormula;   // Update the formula

        // Recalculate all formulas in the workbook
        workbook.CalculateFormula();

        // Verify that the formula has been updated and display the new result
        Console.WriteLine("Updated formula: " + formulaCell.Formula);
        Console.WriteLine("Updated value: " + formulaCell.Value);

        // Save the modified workbook
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}