// Title: C# example: add whole-number data validation to A1:A10 and recalculate dependent formulas with Aspose.Cells
// Common Searches: how to apply integer data validation to a column and recalculate formulas using Aspose.Cells in C# | Aspose.Cells C# example for adding whole number validation to A1:A10 and updating dependent formulas | C# code to set data validation for cells A1 to A10 and trigger Workbook.CalculateFormula in Aspose.Cells | recalculate formulas after adding data validation with Aspose.Cells .NET
// Tags: whole number validation Aspose.Cells C# | Workbook.CalculateFormula recalculate formulas | CellArea range definition Aspose.Cells | apply data validation to range Aspose.Cells | save workbook to xlsx Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new Workbook, defines a CellArea covering A1:A10, adds whole-number validation (values 1‑100) to that range, calls Workbook.CalculateFormula to update any formulas that depend on the validated cells, and saves the result as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // Define the range for validation (A1:A10)
            var area = new CellArea
            {
                StartRow = 0,      // Row 1 (zero‑based)
                StartColumn = 0,   // Column A (zero‑based)
                EndRow = 9,        // Row 10
                EndColumn = 0      // Column A
            };

            // Add integer (whole number) validation to the specified range
            int validationIndex = sheet.Validations.Add(area);
            var validation = sheet.Validations[validationIndex];
            validation.Type = ValidationType.WholeNumber;   // whole numbers between 1 and 100
            validation.Operator = OperatorType.Between;
            validation.Formula1 = "1";
            validation.Formula2 = "100";

            // Recalculate formulas that depend on the validated cells
            workbook.CalculateFormula();

            // Save the workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
