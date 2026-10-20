// Title: Set workbook calculation mode to Automatic and retrieve a single cell formula result with Aspose.Cells for .NET (C#)
// Common Searches: how to enable automatic calculation mode in Aspose.Cells C# | get result of a formula in a specific Excel cell using Aspose.Cells .NET | read evaluated value from cell after calling CalculateFormula in C# | programmatically calculate workbook formulas with Aspose.Cells | example of setting calculation mode and retrieving cell value in Aspose.Cells
// Tags: automatic calculation mode Aspose.Cells | calculate workbook formulas .NET | retrieve evaluated cell value C# | set calculation mode programmatically Aspose.Cells | read cell value after formula evaluation Aspose.Cells

using Aspose.Cells;
using System;

// The sample creates a new Workbook (default Automatic calculation mode), writes a SUM formula to cell A1, triggers full workbook calculation with CalculateFormula, reads the evaluated result from the cell's Value property, and outputs the result to the console, all wrapped in exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Put a formula in cell A1 (example: sum of numbers)
            Cell targetCell = sheet.Cells["A1"];
            targetCell.Formula = "=SUM(10,20,30)";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Retrieve the evaluated result from the cell
            object result = targetCell.Value;

            // Output the evaluated result
            Console.WriteLine($"Result of formula in A1: {result}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
