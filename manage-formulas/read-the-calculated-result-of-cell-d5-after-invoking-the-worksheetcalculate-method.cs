// Title: Read the calculated value of cell D5 after evaluating formulas with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get value of a cell after Workbook.CalculateFormula | retrieve D5 result after worksheet calculation using Aspose.Cells .NET | how to read evaluated cell value in Excel file with Aspose.Cells C# | C# Aspose.Cells read formula result from specific cell after recalculation | display calculated value of D5 after calling CalculateFormula in Aspose.Cells
// Tags: Workbook.CalculateFormula retrieve cell value | Aspose.Cells read calculated cell | C# get evaluated Excel cell value | worksheet calculation result extraction Aspose.Cells | read cell D5 after formula evaluation

using System;
using System.IO;
using Aspose.Cells;

// The program loads an existing Excel workbook, triggers full formula recalculation with Workbook.CalculateFormula, accesses the first worksheet, reads the value of cell D5 after calculation, converts it to a string, and writes the result to the console.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Get the first worksheet (or specify by name/index)
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve cell D5 after calculation
            Cell cell = worksheet.Cells["D5"];
            object result = cell.Value; // Calculated result

            // Convert to string for display (adjust type as needed)
            string resultString = result?.ToString() ?? string.Empty;

            // Output the result
            Console.WriteLine($"Calculated value of D5: {resultString}");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
