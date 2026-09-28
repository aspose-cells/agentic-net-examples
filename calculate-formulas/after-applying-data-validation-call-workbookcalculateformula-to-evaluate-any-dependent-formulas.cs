// Title: Add whole-number data validation referencing another cell and recalculate dependent formulas with Aspose.Cells for .NET
// AI Prompts: Create a workbook, set a whole-number validation on A1 with its maximum value linked to B1, then call Workbook.CalculateFormula to refresh formulas. | Define a validation rule that uses a cell reference, assign a formula that depends on the validated cell, and invoke CalculateFormula to evaluate it in C#.
// Common Searches: Aspose.Cells how to use a cell reference as the upper limit in data validation | C# recalculate formulas after adding data validation with Aspose.Cells | example of Workbook.CalculateFormula with dependent formulas in .NET | set whole number validation between a constant and another cell using Aspose.Cells | update formula results after changing validation range in Aspose.Cells workbook
// Tags: Aspose.Cells validation range from cell reference | Workbook.CalculateFormula example in C# | evaluate dependent formulas after validation Aspose.Cells | dynamic upper bound validation B1 Aspose.Cells | C# data validation and formula recalculation Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsValidationCalc
{
    // Shows how to create a workbook, add a whole-number validation to A1 whose maximum value is taken from B1, set a formula in C1 that depends on A1, invoke Workbook.CalculateFormula to evaluate the formula, and save the file as an XLSX document.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and its cells
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Put a reference value in B1 (used by validation)
            cells["B1"].PutValue(50);

            // Add data validation to cell A1 (whole number between 10 and the value in B1)
            Validation validation = sheet.Validations[sheet.Validations.Add()];
            validation.Type = ValidationType.WholeNumber;
            validation.Operator = OperatorType.Between;
            validation.Formula1 = "10";   // Minimum value
            validation.Formula2 = "=B1"; // Maximum value referencing B1

            // Define the area (A1) to which the validation applies
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 0,
                EndColumn = 0
            };
            validation.AddArea(area);

            // Set a formula that depends on the validated cell
            cells["C1"].Formula = "=A1*2";

            // Calculate all formulas after applying validation
            workbook.CalculateFormula();

            // Save the workbook
            workbook.Save("ValidationAndCalc.xlsx", SaveFormat.Xlsx);
        }
    }
}
