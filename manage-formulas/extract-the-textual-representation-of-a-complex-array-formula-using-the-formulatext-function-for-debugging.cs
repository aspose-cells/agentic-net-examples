// Title: Extract the textual representation of a complex array formula with FORMULATEXT using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get formula string from a cell using FORMULATEXT | debugging array formulas in .NET with Aspose.Cells FORMULATEXT | retrieve Excel array formula text programmatically with Aspose.Cells | how to display a complex formula as text in an Aspose.Cells workbook | extract formula text for debugging in C# Aspose.Cells example
// Tags: FORMULATEXT usage in Aspose.Cells .NET | extracting array formula text C# | Aspose.Cells workbook formula evaluation | displaying Excel formula as string | C# console output of formula text

using Aspose.Cells;
using System;

// The example creates a workbook, assigns a complex array formula to cell A1, uses the FORMULATEXT function in cell B1 to capture the formula's textual representation, calculates all formulas, reads the resulting string from B1, and writes the extracted formula text to the console for debugging purposes.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define a complex array formula (example)
            string arrayFormula = "SUM(IF({1,2,3}>1,{1,2,3},0))";

            // Apply the formula to cell A1 (Aspose.Cells treats this as an array formula when appropriate)
            sheet.Cells["A1"].Formula = arrayFormula;

            // Retrieve the formula text of A1 into B1 using FORMULATEXT
            sheet.Cells["B1"].Formula = "=FORMULATEXT(A1)";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Read the textual representation from B1
            string formulaText = sheet.Cells["B1"].StringValue;

            // Output the result for debugging
            Console.WriteLine("Array formula text: " + formulaText);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
