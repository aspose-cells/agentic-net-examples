// Title: Use Aspose.Cells for .NET to calculate all workbook formulas and fetch the evaluated value of cell G10
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, executes Workbook.CalculateFormula(), and returns the Value of cell G10 from the first worksheet. | Create a console application that evaluates every formula in a workbook using Aspose.Cells and prints the computed result of cell G10.
// Common Searches: Aspose.Cells C# calculate all formulas and get value of a specific cell | How to retrieve evaluated result of cell G10 after Workbook.CalculateFormula in .NET | C# example for evaluating Excel formulas with Aspose.Cells and reading a cell value | Get calculated value of a cell after formula evaluation using Aspose.Cells for .NET
// Tags: Aspose.Cells evaluate workbook formulas | retrieve evaluated cell value C# | Workbook.CalculateFormula example | read cell G10 Aspose.Cells | Excel formula calculation .NET

using System;
using Aspose.Cells;

// The sample loads an Excel workbook with Aspose.Cells, triggers calculation of all formulas via Workbook.CalculateFormula(), accesses cell G10 on the first worksheet, extracts its evaluated value, and writes the result to the console.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Evaluate all formulas in the workbook
        workbook.CalculateFormula();

        // Access cell G10 from the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cell targetCell = sheet.Cells["G10"];

        // Retrieve the calculated value (as an object)
        object calculatedValue = targetCell.Value;

        // Output the value to the console
        Console.WriteLine(calculatedValue);
    }
}
