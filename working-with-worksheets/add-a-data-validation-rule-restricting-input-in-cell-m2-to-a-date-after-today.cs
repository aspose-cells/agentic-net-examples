// Title: Add Future‑Date Validation to Cell M2 with Aspose.Cells for .NET (C#)
// Description: Learn how to use Aspose.Cells for .NET to create a workbook, target cell M2, and apply a data‑validation rule that only accepts dates later than TODAY(). Includes custom input and error messages.
// Keywords: Aspose.Cells | C# data validation | Excel future date validation | cell M2 validation | date after today | ValidationType.Date | OperatorType.GreaterThan | TODAY() formula | .NET Excel automation
// Common Searches: Aspose.Cells set date validation after today | C# add future date rule to Excel cell | How to restrict Excel cell to future dates using Aspose | Aspose.Cells validation for cell M2 | Excel data validation with TODAY() in .NET
// Developer Intent: Implement a validation rule that permits only dates later than the current day in cell M2 of an Excel workbook using Aspose.Cells for C#.
// Use Cases: Ensure scheduling worksheets accept only upcoming appointment dates. | Prevent entry of past due dates in project timelines. | Validate delivery dates in purchase order forms, showing custom prompts.
// AI Prompts: Generate C# code with Aspose.Cells that adds a data‑validation rule to cell M2 allowing only dates after TODAY(), with custom input and error messages. | Show how to modify the rule to reference a user‑defined start date instead of TODAY() and apply it to a range of cells. | Provide an example that reads existing validation settings from a workbook and updates the formula to point to a cell containing the current date.

using System;
using Aspose.Cells;

// Learn how to use Aspose.Cells for .NET to create a workbook, target cell M2, and apply a data‑validation rule that only accepts dates later than TODAY(). Includes custom input and error messages.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define the cell area for M2 (row index 1, column index 12)
        CellArea area = CellArea.CreateCellArea(1, 12, 1, 12);

        // Add a new validation to the worksheet for the defined area
        int validationIndex = sheet.Validations.Add(area);
        Validation validation = sheet.Validations[validationIndex];

        // Configure the validation to allow only dates greater than TODAY()
        validation.Type = ValidationType.Date;                 // Date validation
        validation.Operator = OperatorType.GreaterThan;        // Must be greater than
        validation.Formula1 = "TODAY()";                       // Reference date (today)

        // Optional user interface settings
        validation.ShowInput = true;
        validation.InputTitle = "Date Validation";
        validation.InputMessage = "Please enter a date after today.";
        validation.ShowError = true;
        validation.ErrorTitle = "Invalid Date";
        validation.ErrorMessage = "The date must be after today.";

        // Save the workbook to a file
        workbook.Save("DateValidationAfterToday.xlsx");
    }
}
