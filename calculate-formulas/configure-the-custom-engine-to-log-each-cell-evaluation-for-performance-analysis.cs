// Title: C# example: Log each cell’s formula evaluation with Aspose.Cells CalculationOptions for performance profiling
// Common Searches: how to log each cell formula evaluation using Aspose.Cells in C# | Aspose.Cells calculate cell formulas with custom logging for performance analysis | C# iterate worksheets and log formula results with Aspose.Cells CalculationOptions | profile Excel formula calculation time per cell using Aspose.Cells | track cell calculation values while saving workbook with Aspose.Cells C#
// Tags: Aspose.Cells calculation logging | C# CalculationOptions performance analysis | Worksheet iteration formula evaluation Aspose.Cells | Excel calculation profiling Aspose.Cells | Save workbook after formula evaluation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, walks through every worksheet and each used cell, logs the cell address and its formula before calling Calculate with a CalculationOptions instance, logs the resulting value, and finally saves the workbook, handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Prepare calculation options (default options are sufficient for most cases)
            CalculationOptions calcOptions = new CalculationOptions();

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all cells in the used range of the worksheet
                foreach (Cell cell in sheet.Cells)
                {
                    // Process only cells that contain a formula
                    if (cell.IsFormula)
                    {
                        // Log the cell address and its formula before evaluation
                        Console.WriteLine($"Evaluating cell {cell.Name} - Formula: {cell.Formula}");

                        // Evaluate the formula for this cell using the specified options
                        cell.Calculate(calcOptions);

                        // Log the result after evaluation
                        Console.WriteLine($"Result for {cell.Name}: {cell.Value}");
                    }
                }
            }

            // Save the workbook after evaluation
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
