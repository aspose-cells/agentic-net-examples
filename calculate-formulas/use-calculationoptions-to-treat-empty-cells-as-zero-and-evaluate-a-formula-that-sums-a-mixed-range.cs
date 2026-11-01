// Title: Use Aspose.Cells CalculationOptions in C# to treat blank cells as zero when evaluating a SUM formula over a mixed range
// Common Searches: Aspose.Cells C# calculate SUM formula with empty cells treated as zero | How to set CalculationOptions to treat blanks as zero in Aspose.Cells workbook | C# sum range A1:B2 including blank cells using Aspose.Cells | Evaluate Excel formula with missing values using Aspose.Cells CalculationOptions | Aspose.Cells calculate formula ignoring empty cells in C#
// Tags: Aspose.Cells CalculationOptions zero handling | C# SUM formula blank cells | evaluate mixed range formula Aspose.Cells | workbook.CalculateFormula empty cells | Aspose.Cells treat blanks as zero

using System;
using Aspose.Cells;

// The program creates a new workbook, fills cells A1, A2, and B2 with numbers while leaving B1 empty, configures CalculationOptions to treat empty cells as zero, assigns the SUM(A1:B2) formula to C1, calculates all formulas, and prints the resulting sum to the console.
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

            // Populate cells with numbers and leave some empty
            sheet.Cells["A1"].PutValue(10);   // A1 = 10
            // B1 is left empty
            sheet.Cells["A2"].PutValue(20);   // A2 = 20
            sheet.Cells["B2"].PutValue(30);   // B2 = 30

            // Insert the SUM formula into a cell
            sheet.Cells["C1"].Formula = "SUM(A1:B2)";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Retrieve the calculated result
            double sum = sheet.Cells["C1"].DoubleValue;

            // Output the result
            Console.WriteLine("Sum treating empty cells as zero: " + sum);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
