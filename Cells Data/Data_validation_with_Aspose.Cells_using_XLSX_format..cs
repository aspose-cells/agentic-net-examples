using System;
using Aspose.Cells;

class DataValidationExample
{
    static void Main()
    {
        // Paths for the template and the output workbook
        string templatePath = "template.xlsx";
        string outputPath = "output.xlsx";

        // Create load options for XLSX format and disable data‑validation checking
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.CheckDataValid = false; // use LoadOptions.CheckDataValid property

        // Load the workbook with the specified options
        Workbook workbook = new Workbook(templatePath, loadOptions);
        Worksheet sheet = workbook.Worksheets[0];

        // Add a whole‑number validation (between 10 and 100) to cells A1:A10
        Validation validation = sheet.Validations[sheet.Validations.Add()];
        validation.Type = ValidationType.WholeNumber;      // set validation type
        validation.Operator = OperatorType.Between;        // set operator
        validation.Formula1 = "10";                        // lower bound
        validation.Formula2 = "100";                       // upper bound

        // Define the cell area for the validation (A1:A10)
        CellArea area = new CellArea
        {
            StartRow = 0,    // row 1 (zero‑based)
            StartColumn = 0, // column A
            EndRow = 9,      // row 10
            EndColumn = 0    // column A
        };
        validation.AddArea(area); // apply validation to the area

        // Example: retrieve the first formula in A1 notation
        string formula1 = validation.GetFormula1(false, false);
        Console.WriteLine("Formula1 (A1 notation): " + formula1);

        // Save the modified workbook
        workbook.Save(outputPath);
    }
}