using System;
using Aspose.Cells;

class RetrieveFormula
{
    static void Main()
    {
        // Load the workbook from a file
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Access the second worksheet (index 1)
        Worksheet secondSheet = workbook.Worksheets[1];

        // Retrieve the formula from cell E3 before any deletion occurs
        string formulaBeforeDeletion = secondSheet.Cells["E3"].Formula;

        Console.WriteLine("Formula in Sheet2!E3 before deletion: " + formulaBeforeDeletion);

        // (Optional) Example of deleting the cell after capturing the formula
        // secondSheet.Cells.Remove(2, 4); // Row 2 (E3), Column 4 (E)

        // (Optional) Save the workbook if further changes are made
        // workbook.Save("output.xlsx");
    }
}