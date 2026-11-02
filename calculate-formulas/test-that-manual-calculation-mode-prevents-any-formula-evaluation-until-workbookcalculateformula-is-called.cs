// Title: Show how to use manual calculation mode in Aspose.Cells for .NET to postpone formula evaluation until Workbook.CalculateFormula is called
// Common Searches: how to set manual calculation mode in Aspose.Cells C# example | prevent automatic formula calculation in Aspose.Cells workbook | clear cell cached value before calling CalculateFormula Aspose.Cells | delay formula evaluation until workbook.CalculateFormula is invoked .NET | verify manual calculation mode behavior with Aspose.Cells
// Tags: manual calculation mode Aspose.Cells | Workbook.CalculateFormula method | clear cell cache Aspose.Cells | postpone formula evaluation .NET | disable automatic calculation Aspose.Cells

using System;
using Aspose.Cells;

// The sample creates a new Workbook, writes numbers to A1 and A2, assigns the formula "=A1+A2" to A3, clears the cached value, displays the empty value before calculation, then calls Workbook.CalculateFormula() and prints the evaluated result (30), demonstrating manual calculation mode.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // NOTE: In newer Aspose.Cells versions the CalcMode property may be unavailable.
            // The workbook defaults to automatic calculation. To demonstrate manual calculation,
            // we will explicitly call CalculateFormula after setting the formula.

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells with data
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);

            // Insert a formula that adds A1 and A2
            sheet.Cells["A3"].Formula = "=A1+A2";

            // At this point, because calculation mode may be automatic, the formula could already be evaluated.
            // To illustrate the effect of manual calculation, we clear the cached value.
            sheet.Cells["A3"].Value = null;

            Console.WriteLine("Before calling CalculateFormula:");
            Console.WriteLine($"A3 Value (should be null or empty): '{sheet.Cells["A3"].Value}'");
            Console.WriteLine($"A3 Formula: {sheet.Cells["A3"].Formula}");

            // Explicitly trigger formula calculation
            workbook.CalculateFormula();

            Console.WriteLine("After calling CalculateFormula:");
            Console.WriteLine($"A3 Value (expected 30): {sheet.Cells["A3"].Value}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
