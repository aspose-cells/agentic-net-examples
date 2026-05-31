using Aspose.Cells;
using System;

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

        // Add a validation to the defined area and retrieve it
        int validationIndex = worksheet.Validations.Add(area);
        Validation validation = worksheet.Validations[validationIndex];

        // Set the validation type to Custom and set Formula1 to reference A1 using A1 notation
        validation.Type = ValidationType.Custom;
        validation.SetFormula1("=A1", false, false);

        // Get Formula1 back in A1 notation
        string formula1 = validation.GetFormula1(false, false);

        // Verify that the formula references the first worksheet cell A1
        bool referencesA1 = formula1.Contains("A1");
        Console.WriteLine("Formula1 retrieved: " + formula1);
        Console.WriteLine("References A1: " + referencesA1);

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("ValidateFormula.xlsx");
    }
}