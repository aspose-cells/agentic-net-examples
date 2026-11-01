// Title: Compare original and modified cell formula text in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells compare original and new formula text in a cell | How to detect if an Excel cell formula was changed after editing with Aspose.Cells | Replace SUM function with AVERAGE in a formula using Aspose.Cells .NET | Retrieve cell formula string before modification in Aspose.Cells C# | Save workbook after updating cell formula with Aspose.Cells in C#
// Tags: Aspose.Cells read formula string C# | Aspose.Cells replace function in Excel formula | Aspose.Cells compare original and updated formulas | Aspose.Cells detect formula modification | Aspose.Cells save workbook after formula edit | Aspose.Cells load or create workbook C#

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook (or creates a new one), reads the formula in cell A1, replaces the SUM function with AVERAGE, captures both the original and modified formula strings, compares them to determine if a change occurred, outputs the results, and saves the workbook with the updated formula.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                // Load without password (assumes file is not password‑protected).
                workbook = new Workbook(inputPath);
            }
            else
            {
                // Create a new workbook with a default worksheet.
                workbook = new Workbook();
                Worksheet defaultSheet = workbook.Worksheets[0];
                // Set a sample formula for demonstration.
                defaultSheet.Cells["A1"].Formula = "SUM(B1:B5)";
            }

            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];

            // Capture the original formula text.
            string originalFormula = cell.Formula;

            // Example modification: replace SUM with AVERAGE.
            if (!string.IsNullOrEmpty(originalFormula))
            {
                string modifiedFormula = originalFormula.Replace("SUM", "AVERAGE");
                cell.Formula = modifiedFormula;
            }

            // Capture the new formula text after modification.
            string newFormula = cell.Formula;

            // Compare before and after formula texts.
            bool formulaChanged = !string.Equals(originalFormula, newFormula, StringComparison.Ordinal);

            // Output the comparison results.
            Console.WriteLine($"Original formula: {originalFormula}");
            Console.WriteLine($"New formula: {newFormula}");
            Console.WriteLine($"Formula changed: {formulaChanged}");

            // Save the workbook with the updated formula.
            workbook.Save(outputPath);
        }
        catch (CellsException ex)
        {
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
