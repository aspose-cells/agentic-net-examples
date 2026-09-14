// Title: Calculate all formulas in an Excel workbook and obtain the evaluated value of cell G10 with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file using Aspose.Cells, runs workbook.CalculateFormula(), and prints the computed value of cell G10. | Write a reusable C# method that takes a workbook path and a cell address, evaluates every formula with Aspose.Cells, and returns the resulting cell value. | Show how to access the first worksheet of a loaded workbook and read the Value property of a cell after calling CalculateFormula in Aspose.Cells.
// Common Searches: asp.net calculate all formulas in Excel file using Aspose.Cells and read cell G10 result | C# Aspose.Cells workbook.CalculateFormula example for retrieving a specific cell value | how to get evaluated value of a formula cell after calling CalculateFormula with Aspose.Cells | retrieve value of G10 from first worksheet after formula evaluation using Aspose.Cells .NET | Aspose.Cells read calculated result of a cell without opening Excel
// Tags: calculate formulas Aspose.Cells .NET | read evaluated cell value C# | access first worksheet Aspose.Cells | retrieve cell G10 after calculation | load workbook input.xlsx Aspose.Cells

using Aspose.Cells;
using System;

// The program loads 'input.xlsx' with Aspose.Cells, evaluates all formulas in the workbook, accesses cell G10 on the first worksheet, and outputs its calculated value to the console.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Evaluate all formulas in the workbook
        workbook.CalculateFormula();

        // Access the first worksheet (index 0)
        Worksheet worksheet = workbook.Worksheets[0];

        // Retrieve the calculated value of cell G10
        Cell cellG10 = worksheet.Cells["G10"];
        object calculatedValue = cellG10.Value;

        // Output the result
        Console.WriteLine("Calculated value of G10: " + calculatedValue);
    }
}
