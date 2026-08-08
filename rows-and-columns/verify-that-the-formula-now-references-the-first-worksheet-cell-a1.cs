// Title: Check that a custom data‑validation formula points to cell A1 using Aspose.Cells for .NET
// Description: Creates a new workbook, adds a custom validation to cell B1, sets its Formula1 to "=A1" with A1 notation, retrieves the formula, confirms it references A1, prints the result, and saves the file as VerifyFormulaReference.xlsx.
// Keywords: Aspose.Cells | .NET | C# | data validation | custom validation | Formula1 | A1 notation | verify formula reference | cell reference validation | Excel automation
// Common Searches: Aspose.Cells set validation formula to A1 | how to retrieve validation formula in Aspose.Cells C# | verify data validation cell reference .NET | custom validation example Aspose.Cells | check if validation formula points to specific cell
// Developer Intent: Confirm that the custom validation applied to cell B1 correctly references cell A1.
// Use Cases: Programmatically add a custom validation rule and ensure it targets the intended source cell before publishing the workbook. | Automated testing of Excel files to validate that data‑validation formulas reference the correct cells across multiple sheets. | Generate reports that list validation formulas and flag any that do not match expected cell addresses.
// AI Prompts: Write C# code with Aspose.Cells to add a custom validation to cell C3 that references cell D5 and verify the reference. | Create a unit test in C# that asserts a validation formula in a workbook equals a given cell address using Aspose.Cells. | Explain how to retrieve a validation formula in A1 notation and compare it to an expected reference in Aspose.Cells.

using Aspose.Cells;
using System;

// Creates a new workbook, adds a custom validation to cell B1, sets its Formula1 to "=A1" with A1 notation, retrieves the formula, confirms it references A1, prints the result, and saves the file as VerifyFormulaReference.xlsx.
class VerifyFormulaReference
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the area for the validation (cell B1)
        CellArea area = new CellArea
        {
            StartRow = 0,
            StartColumn = 1,
            EndRow = 0,
            EndColumn = 1
        };

        // Add the validation to the worksheet
        int validationIndex = worksheet.Validations.Add(area);
        Validation validation = worksheet.Validations[validationIndex];
        validation.Type = ValidationType.Custom;

        // Set Formula1 to reference the first cell A1 using A1 notation
        validation.SetFormula1("=A1", false, false);

        // Retrieve the formula in A1 notation
        string retrievedFormula = validation.GetFormula1(false, false);

        // Verify that the formula references cell A1
        bool referencesA1 = retrievedFormula.TrimStart('=') == "A1";

        Console.WriteLine("Retrieved Formula: " + retrievedFormula);
        Console.WriteLine("References A1: " + referencesA1);

        // Save the workbook (optional verification step)
        workbook.Save("VerifyFormulaReference.xlsx");
    }
}
