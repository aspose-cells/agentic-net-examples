// Title: How to read the evaluated value of cell D5 after calling Workbook.CalculateFormula with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get the result of a formula in cell D5 after Workbook.CalculateFormula | retrieve evaluated cell value after worksheet.Calculate in Aspose.Cells | Aspose.Cells example to read calculated value from Excel cell D5 in .NET | how to obtain formula output from a specific cell after recalculating workbook using Aspose.Cells C# | read cell D5 value after invoking Workbook.CalculateFormula with Aspose.Cells
// Tags: Workbook.CalculateFormula retrieve cell value | Aspose.Cells read evaluated cell D5 | C# get formula result from Excel cell | Aspose.Cells calculate all formulas and access cell | Excel cell value after recalculation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing workbook, triggers full formula recalculation with Workbook.CalculateFormula(), accesses the first worksheet, reads the evaluated value of cell D5 via the Cells collection, and prints the result to the console while handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Access the first worksheet (or specify by name/index)
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the calculated value of cell D5
            Cell cellD5 = worksheet.Cells["D5"];
            object calculatedValue = cellD5.Value; // could be double, string, DateTime, etc.

            // Output the result
            Console.WriteLine($"Calculated value of D5: {calculatedValue}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
