// Title: Extract array formula text using FORMULATEXT with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to get the text of an array formula | using FORMULATEXT in Aspose.Cells to debug Excel formulas | retrieve Excel formula string after calculation with Aspose.Cells .NET | Aspose.Cells extract SUM(A1:A5*B1:B5) formula text programmatically | debugging array formulas in C# by reading FORMULATEXT result
// Tags: Aspose.Cells extract formula text | FORMULATEXT array formula debugging | C# read Excel formula text | calculate workbook before FORMULATEXT Aspose.Cells | Excel array formula extraction .NET

using Aspose.Cells;
using System;
using System.IO;

// The example creates a workbook, populates columns A and B, assigns an array formula (SUM(A1:A5*B1:B5)) to cell C1, uses the Excel FORMULATEXT function in D1 to capture the formula as a string, calculates the workbook so FORMULATEXT evaluates, reads the resulting text from D1 for debugging, prints it, and saves the file as ArrayFormulaDebug.xlsx.
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

            // Fill sample data in columns A and B (A1:A5 and B1:B5)
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[i, 0].PutValue(i + 1);          // Column A
                sheet.Cells[i, 1].PutValue((i + 1) * 2);    // Column B
            }

            // Set an array formula in C1. Aspose.Cells supports FormulaArray, but if unavailable use regular formula.
            // This will still allow FORMULATEXT to retrieve the formula text.
            sheet.Cells["C1"].Formula = "SUM(A1:A5*B1:B5)";

            // Use the Excel FORMULATEXT function to retrieve the formula text of C1 into D1
            sheet.Cells["D1"].Formula = "FORMULATEXT(C1)";

            // Calculate the workbook so FORMULATEXT is evaluated
            workbook.CalculateFormula();

            // Extract the textual representation of the array formula from D1
            string formulaText = sheet.Cells["D1"].StringValue;

            // Output the result for debugging
            Console.WriteLine("Array formula in C1: " + formulaText);

            // Define output file path
            string outputPath = "ArrayFormulaDebug.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
