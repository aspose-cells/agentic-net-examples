// Title: Add list data validation linked to the named range "ValidCodes" in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: how to add a dropdown list validation from a named range with Aspose.Cells C# | Aspose.Cells C# list validation using existing named range | set data validation formula to a named range in Aspose.Cells workbook | apply list‑type validation to a cell range using Aspose.Cells .NET | create Excel data validation referencing named range ValidCodes with Aspose.Cells
// Tags: list validation using named range Aspose.Cells | Aspose.Cells dropdown validation C# | Excel validation formula referencing named range | cell area validation Aspose.Cells .NET | data validation error alert configuration Aspose

using Aspose.Cells;
using System;

// The example creates a workbook, defines the cell area C1:C10, adds a List‑type validation that points to the existing named range "ValidCodes", enables an in‑cell drop‑down with an error alert, and saves the file as DataValidationExample.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Assume the named range "ValidCodes" already exists in the workbook.
        // If you need to create it, you could do something like:
        // Range validCodes = sheet.Cells.CreateRange("B1:B5");
        // validCodes.Name = "ValidCodes";

        // Define the area where the validation will be applied (e.g., C1:C10)
        CellArea validationArea = new CellArea
        {
            StartRow = 0,
            EndRow = 9,
            StartColumn = 2,
            EndColumn = 2
        };

        // Add a List-type data validation that references the named range "ValidCodes"
        int validationIndex = sheet.Validations.Add(validationArea);
        Validation validation = sheet.Validations[validationIndex];
        validation.Type = ValidationType.List;
        validation.Operator = OperatorType.Between; // Required placeholder for List type
        validation.Formula1 = "=ValidCodes"; // Reference to the named range
        validation.InCellDropDown = true;
        validation.ShowError = true;
        validation.ErrorTitle = "Invalid Code";
        validation.ErrorMessage = "Please select a value from the list of valid codes.";

        // Save the workbook
        workbook.Save("DataValidationExample.xlsx");
    }
}
