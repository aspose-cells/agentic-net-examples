// Title: C# Aspose.Cells: Add Date‑After‑Today Validation to Cell M2
// Description: Shows how to create a workbook with Aspose.Cells for .NET, target cell M2, and apply a data‑validation rule that accepts only dates later than the current day. Includes optional input and error messages and saves the file as DateValidationAfterToday.xlsx.
// Keywords: Aspose.Cells | C# | Excel data validation | date after today | cell M2 | validation type date | OperatorType.GreaterThan | TODAY() formula | worksheet validation | Excel automation
// Common Searches: Aspose.Cells set date validation after today | C# add data validation to specific Excel cell | restrict Excel cell to future dates using Aspose | validate cell M2 with Aspose.Cells .NET | Excel date validation formula TODAY() Aspose
// Developer Intent: Add a data‑validation rule that permits only dates later than today in cell M2 of an Excel worksheet using Aspose.Cells for .NET.
// Use Cases: Ensuring future appointment dates in scheduling templates. | Preventing past due dates in invoice or billing sheets. | Requiring delivery dates after the current day in logistics workbooks. | Guiding users to enter upcoming event dates in project plans.
// AI Prompts: Generate C# code with Aspose.Cells that enforces a date‑after‑today rule for cell M2. | Create a reusable method to apply a future‑date validation to any cell address using Aspose.Cells. | Explain how to customize the input and error dialogs for a date validation rule in Aspose.Cells.

using Aspose.Cells;
using System;

// Shows how to create a workbook with Aspose.Cells for .NET, target cell M2, and apply a data‑validation rule that accepts only dates later than the current day. Includes optional input and error messages and saves the file as DateValidationAfterToday.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define the cell area for M2 (row 1, column 12 in zero‑based indexes)
        CellArea area = CellArea.CreateCellArea(1, 12, 1, 12);

        // Add a validation to the worksheet for the defined area
        int validationIndex = sheet.Validations.Add(area);
        Validation validation = sheet.Validations[validationIndex];

        // Configure the validation: Date type, must be greater than TODAY()
        validation.Type = ValidationType.Date;
        validation.Operator = OperatorType.GreaterThan;
        validation.Formula1 = "=TODAY()";

        // Optional: display input message when the cell is selected
        validation.ShowInput = true;
        validation.InputTitle = "Date Validation";
        validation.InputMessage = "Enter a date after today.";

        // Optional: display error message for invalid input
        validation.ShowError = true;
        validation.ErrorTitle = "Invalid Date";
        validation.ErrorMessage = "The date must be later than today.";

        // Save the workbook
        workbook.Save("DateValidationAfterToday.xlsx");
    }
}
