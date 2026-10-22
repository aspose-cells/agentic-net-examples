// Title: Check that a cell’s string length stays the same after Workbook.CalculateFormula with Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells to write a text string to a worksheet cell, call Workbook.CalculateFormula, and assert that the original string length equals the length after recalculation. | Create a reusable C# method that accepts a cell reference, records its text length, triggers workbook calculation, and returns a boolean indicating whether the length was preserved. | Adapt an Aspose.Cells example to log the original and post‑calculation string lengths and output a pass/fail message based on the length comparison.
// Common Searches: aspnet cells calculateformula keep cell text length unchanged | c# verify excel cell string length after workbook refresh using Aspose.Cells | how to assert cell content length consistency after recalculating formulas Aspose.Cells | Aspose.Cells workbook recalc does not truncate cell string length
// Tags: Aspose.Cells CalculateFormula preserve cell string length | C# verify cell text length after workbook refresh | Aspose.Cells workbook recalculation length consistency | Excel cell length validation with Aspose.Cells | Debug.Assert cell length unchanged Aspose.Cells

using System;
using System.Diagnostics;
using Aspose.Cells;

// The example creates a new workbook, writes a specific text string to cell A1, records its length, calls Workbook.CalculateFormula to refresh the workbook, then compares the original and post‑refresh string lengths using Debug.Assert, prints the results, and saves the workbook as VerificationResult.xlsx.
class VerifyCellContentLength
{
    static void Main()
    {
        try
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set a string value in cell A1
            string originalText = "This is a test string with a specific length.";
            cells["A1"].PutValue(originalText);

            // Record the original length of the cell content
            int originalLength = cells["A1"].StringValue.Length;

            // Refresh the workbook (recalculate formulas, etc.)
            workbook.CalculateFormula();

            // Retrieve the cell content after refresh
            string afterText = cells["A1"].StringValue;
            int afterLength = afterText.Length;

            // Verify that the length remains unchanged
            Debug.Assert(originalLength == afterLength, "Cell content length changed after refresh.");

            // Output the verification result
            Console.WriteLine("Original Length: {0}", originalLength);
            Console.WriteLine("Length After Refresh: {0}", afterLength);
            Console.WriteLine(originalLength == afterLength
                ? "Verification passed: Length unchanged."
                : "Verification failed: Length changed.");

            // Save the workbook (save rule)
            string outputPath = "VerificationResult.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to: " + outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
