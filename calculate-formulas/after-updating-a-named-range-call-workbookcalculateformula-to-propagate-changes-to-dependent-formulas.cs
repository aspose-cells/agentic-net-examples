// Title: Update a named range and recalculate dependent formulas using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells change defined name reference and refresh formulas | Update Excel named range programmatically and recalculate formulas with Aspose.Cells | Calling Workbook.CalculateFormula after modifying a named range in .NET | Aspose.Cells example for renaming a range and propagating formula changes | How to adjust MyRange reference and recalc dependent cells using Aspose.Cells
// Tags: update named range Aspose.Cells | recalculate dependent formulas Workbook.CalculateFormula | modify defined name reference C# | propagate named range changes Excel .NET | Workbook.CalculateFormula method Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, retrieves the defined name "MyRange", updates its RefersTo reference to a new range, invokes Workbook.CalculateFormula to refresh all formulas that depend on the named range, and saves the result, with error handling for missing files or undefined names.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Access the collection of defined names (named ranges)
            NameCollection names = workbook.Worksheets.Names;

            // Retrieve the specific named range
            Name namedRange = names["MyRange"];
            if (namedRange == null)
                throw new InvalidOperationException("Named range 'MyRange' does not exist.");

            // Update the reference of the named range
            namedRange.RefersTo = "Sheet1!C1:D2";

            // Recalculate all formulas that depend on the named range
            workbook.CalculateFormula();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
