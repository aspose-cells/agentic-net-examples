// Title: How to evaluate a single cell formula using Cell.Calculate with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells calculate only cell A1 without full workbook recalculation | Use Cell.Calculate method to get result of a formula in a specific cell in .NET | Aspose.Cells evaluate SUM formula in a single cell with CalculationOptions | Retrieve calculated value from a cell after setting a formula programmatically in C# | Partial workbook calculation using Aspose.Cells Cell.Calculate
// Tags: Cell.Calculate single cell evaluation | Aspose.Cells formula calculation C# | partial workbook recalculation Aspose | CalculationOptions usage Aspose.Cells | extract calculated cell value .NET

using System;
using Aspose.Cells;

// The example creates a new Workbook, assigns the formula "=SUM(1,2,3,4,5)" to cell A1, configures default CalculationOptions, calls targetCell.Calculate to compute the formula, reads the resulting value, prints it, and saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (using the create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set a formula in cell A1 (e.g., sum of numbers 1 to 5)
            Cell targetCell = sheet.Cells["A1"];
            targetCell.Formula = "=SUM(1,2,3,4,5)";

            // Prepare calculation options (default settings)
            CalculationOptions calcOptions = new CalculationOptions();

            // Evaluate the formula for this cell
            targetCell.Calculate(calcOptions);
            object result = targetCell.Value; // Retrieve the calculated value

            // Output the result
            Console.WriteLine($"The result of the formula in A1 is: {result}");

            // Optionally, save the workbook (using the save rule)
            string outputPath = "ResultWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
