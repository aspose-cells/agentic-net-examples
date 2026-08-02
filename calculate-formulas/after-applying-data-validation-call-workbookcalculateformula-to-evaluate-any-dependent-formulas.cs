using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Sample data and a dependent formula
        cells["A1"].PutValue(10);               // Base value
        cells["B1"].Formula = "=A1*2";          // Formula depends on A1

        // Add data validation to cell C1 (whole number between 5 and 20)
        Validation validation = sheet.Validations[sheet.Validations.Add()];
        validation.Type = ValidationType.WholeNumber;
        validation.Operator = OperatorType.Between;
        validation.Formula1 = "5";
        validation.Formula2 = "20";

        // Apply the validation to cell C1
        CellArea area = new CellArea
        {
            StartRow = 0,      // Row 0 (A)
            StartColumn = 2,   // Column C
            EndRow = 0,
            EndColumn = 2
        };
        validation.AddArea(area);

        // After applying validation, calculate all formulas (required step)
        workbook.CalculateFormula();

        // Display calculated results
        Console.WriteLine("A1 value: " + cells["A1"].IntValue);
        Console.WriteLine("B1 formula result: " + cells["B1"].IntValue);

        // Save the workbook (lifecycle rule: save)
        workbook.Save("Output.xlsx");
    }
}