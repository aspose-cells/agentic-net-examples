// Title: Read and print the updated formula and calculated value of cell E3 after workbook calculation using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells get cell formula string after Workbook.CalculateFormula | How to output the evaluated value of a specific Excel cell with Aspose.Cells .NET | Retrieve and display cell E3 formula and result after calculation using Aspose.Cells | Aspose.Cells example to read cell formula and calculated value in a console app | Calculate Excel formulas programmatically and read cell result with Aspose.Cells C#
// Tags: Aspose.Cells read evaluated cell formula | Aspose.Cells calculate workbook formulas C# | C# console display Excel cell value Aspose.Cells | Retrieve cell result after Workbook.CalculateFormula | Sum formula evaluation in Aspose.Cells .NET

using System;
using Aspose.Cells;

// The example creates a new workbook, assigns the formula "=SUM(A1:A5)" to cell E3, forces formula calculation with Workbook.CalculateFormula, then reads the cell's Formula property and its evaluated Value, and writes both the updated formula string and the calculated result to the console.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set a sample formula in cell E3 (row 2, column 4)
        // For demonstration, we sum the range A1:A5
        Cell targetCell = sheet.Cells["E3"];
        targetCell.Formula = "=SUM(A1:A5)";

        // Force calculation of all formulas in the workbook
        workbook.CalculateFormula();

        // Retrieve the updated formula (as a string) and its calculated value
        string updatedFormula = targetCell.Formula;
        object calculatedValue = targetCell.Value;

        // Write the results to the console
        Console.WriteLine("Updated Formula in E3: " + updatedFormula);
        Console.WriteLine("Calculated Value of E3: " + calculatedValue);
    }
}
