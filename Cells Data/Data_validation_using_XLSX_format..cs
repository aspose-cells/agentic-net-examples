using System;
using Aspose.Cells;

class DataValidationDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a whole number validation that requires values between 10 and 100
        Validation validation = worksheet.Validations[worksheet.Validations.Add()];
        validation.Type = ValidationType.WholeNumber;
        validation.Operator = OperatorType.Between;
        validation.Formula1 = "10";
        validation.Formula2 = "100";

        // Apply the validation to the range A1:A10
        CellArea area = new CellArea
        {
            StartRow = 0,      // Row 1 (zero‑based)
            StartColumn = 0,   // Column A (zero‑based)
            EndRow = 9,        // Row 10
            EndColumn = 0      // Column A
        };
        validation.AddArea(area);

        // Save the workbook to an XLSX file
        string filePath = "DataValidationDemo.xlsx";
        workbook.Save(filePath);

        // Load the workbook with data‑validation checking disabled
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.CheckDataValid = false; // Skip validation checks while loading
        Workbook loadedWorkbook = new Workbook(filePath, loadOptions);
        Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];

        // Retrieve the validation that was saved
        Validation loadedValidation = loadedWorksheet.Validations[0];

        // Output validation details to verify it was preserved
        Console.WriteLine($"Validation Type: {loadedValidation.Type}");
        Console.WriteLine($"Operator: {loadedValidation.Operator}");
        Console.WriteLine($"Formula1 (lower bound): {loadedValidation.Formula1}");
        Console.WriteLine($"Formula2 (upper bound): {loadedValidation.Formula2}");
    }
}