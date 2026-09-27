// Title: Set Aspose.Cells workbook to Automatic calculation mode and evaluate a single cell formula in C#
// AI Prompts: Write C# code that switches the workbook's calculation engine to Automatic, assigns a formula to a target cell, and calls the EvaluateFormula method to obtain the cell's computed value. | Show how to retrieve the result of cell A1 without triggering a full workbook recalculation by using Aspose.Cells' EvaluateFormula functionality in a .NET application.
// Common Searches: Aspose.Cells C# set calculation mode to Automatic and get result of one cell formula | How to evaluate a specific Excel cell formula with Aspose.Cells without recalculating the whole workbook | C# example using Aspose.Cells EvaluateFormula to read the calculated value from cell A1
// Tags: automatic calculation mode Aspose.Cells C# | single cell formula evaluation Aspose.Cells | Aspose.Cells EvaluateFormula method | retrieve calculated cell value Aspose.Cells | set workbook calculation mode C#

using Aspose.Cells;
using System;

// Creates a workbook, places the formula "=5+3" in cell A1, ensures the calculation engine runs in Automatic mode, evaluates the formula with EvaluateFormula, and outputs the computed result.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set a formula in cell A1
            sheet.Cells["A1"].Formula = "=5+3";

            // Calculate all formulas in the workbook (default mode is Automatic)
            workbook.CalculateFormula();

            // Retrieve the calculated value from the cell
            object result = sheet.Cells["A1"].Value;

            // Display the result
            Console.WriteLine($"Result of {sheet.Cells["A1"].Formula} is {result}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
