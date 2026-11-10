// Title: How to calculate only formulas in a specific range (A1:D20) using Aspose.Cells for .NET
// Common Searches: Aspose.Cells calculate formulas in a selected range C# | Partial workbook calculation example with Aspose.Cells .NET | Evaluate formulas only in cells A1 to D20 using Aspose.Cells | C# code to iterate and calculate formulas in a defined cell area with Aspose.Cells | Skip full workbook recalculation when updating formulas in Aspose.Cells
// Tags: partial formula calculation Aspose.Cells | Cell.Calculate range evaluation C# | Aspose.Cells selective recalculation | evaluate formulas in defined cell area .NET | skip full workbook calculation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads a workbook, defines the cell area A1:D20, iterates through each cell in that range, and calls Cell.Calculate with default CalculationOptions only for cells that contain formulas, then saves the file, avoiding a full workbook recalculation.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load the workbook if the file exists; otherwise create a new workbook.
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Access the first worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the target range A1:D20.
            CellArea targetRange = CellArea.CreateCellArea("A1", "D20");

            // Prepare calculation options (default settings).
            CalculationOptions calcOptions = new CalculationOptions();

            // Iterate through each cell in the defined range.
            for (int row = targetRange.StartRow; row <= targetRange.EndRow; row++)
            {
                for (int col = targetRange.StartColumn; col <= targetRange.EndColumn; col++)
                {
                    Cell cell = worksheet.Cells[row, col];

                    // If the cell contains a formula, calculate only that cell.
                    if (cell.IsFormula)
                    {
                        cell.Calculate(calcOptions); // Evaluate the cell's formula with options
                    }
                }
            }

            // Save the workbook after partial calculation.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
