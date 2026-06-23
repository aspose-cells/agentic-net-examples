using System;
using System.Collections.Generic;
using Aspose.Cells;

class IdentifyDiv0Errors
{
    static void Main()
    {
        // Create a new workbook (replace with new Workbook("input.xlsx") to load an existing file)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Sample data – you can remove or replace this with your own worksheet content
        cells["A1"].PutValue(10);          // Numerator
        cells["B1"].PutValue(0);           // Denominator (zero)
        cells["C1"].Formula = "=A1/B1";    // Will produce #DIV/0! error
        cells["D1"].Formula = "=A1+5";     // Valid formula

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Collect cells whose calculated result is the #DIV/0! error
        List<Cell> div0ErrorCells = new List<Cell>();
        foreach (Cell cell in cells)
        {
            // IsErrorValue indicates the cell result is an error
            // StringValue returns the error text (e.g., "#DIV/0!")
            if (cell.IsErrorValue && cell.StringValue == "#DIV/0!")
            {
                div0ErrorCells.Add(cell);
            }
        }

        // Output the addresses (names) of the error cells
        Console.WriteLine("Cells containing #DIV/0! error:");
        foreach (Cell errCell in div0ErrorCells)
        {
            Console.WriteLine($"{errCell.Name}  Formula: {errCell.Formula}");
        }

        // Save the workbook (optional)
        workbook.Save("output.xlsx");
    }
}