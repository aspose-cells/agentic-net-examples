// Title: Check that a custom data‑validation formula points to cell A1 on the first worksheet using Aspose.Cells for .NET
// AI Prompts: Create a custom validation for cell A1, assign its Formula1 to "=A1" using A1‑style notation, and save the workbook. | Read the validation's Formula1 back in A1 notation, strip any leading '=', compare it to "A1", and output whether the reference is correct.
// Common Searches: Aspose.Cells how to set custom validation formula to reference A1 without R1C1 | C# verify data validation formula points to a specific cell in an Excel workbook | retrieve validation GetFormula1 A1 notation Aspose.Cells .NET | check if validation formula equals A1 after setting it in Aspose.Cells | example code for confirming validation cell reference in Aspose.Cells
// Tags: validation formula assignment Aspose.Cells | GetFormula1 A1 notation .NET | verify validation cell reference C# | non‑R1C1 formula usage Aspose.Cells | data validation reference check Excel

using Aspose.Cells;
using System;

// The sample creates a new workbook, adds a custom validation to cell A1, sets its formula to "=A1" using A1‑style (non‑R1C1) notation, retrieves the formula back in A1 notation, removes any leading '=', compares it to "A1" to confirm the reference, prints the verification result, and saves the workbook.
class VerifyValidationFormula
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the area for the validation (cell A1)
        CellArea area = new CellArea
        {
            StartRow = 0,
            StartColumn = 0,
            EndRow = 0,
            EndColumn = 0
        };

        // Add a custom validation to the defined area
        int validationIndex = worksheet.Validations.Add(area);
        Validation validation = worksheet.Validations[validationIndex];
        validation.Type = ValidationType.Custom;

        // Set the validation formula to reference cell A1 using A1 notation
        // (non‑R1C1, non‑local)
        validation.SetFormula1("=A1", false, false);

        // Retrieve the formula in A1 notation
        string retrievedFormula = validation.GetFormula1(false, false);

        // Remove a leading '=' if present for easier comparison
        if (retrievedFormula.StartsWith("="))
            retrievedFormula = retrievedFormula.Substring(1);

        // Verify that the formula now references the first worksheet cell A1
        bool referencesA1 = string.Equals(retrievedFormula, "A1", StringComparison.OrdinalIgnoreCase);
        Console.WriteLine("Formula references A1: " + referencesA1);

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("VerifyValidationFormula.xlsx");
    }
}
