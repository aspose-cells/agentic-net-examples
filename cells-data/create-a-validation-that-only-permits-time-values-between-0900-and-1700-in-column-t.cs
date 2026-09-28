// Title: Create a time‑range data validation (09:00‑17:00) for column T in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that adds a Validation object to column T (rows 1‑1000) allowing only times from 09:00 to 17:00 and shows custom input and error alerts. | Generate a snippet that changes the validation to a different column index and modifies the lower and upper time bounds while preserving the message settings. | Provide an example of applying the Between operator for time values in Aspose.Cells and exporting the workbook to a file.
// Common Searches: aspocells c# how to restrict a column to business hours 09:00-17:00 | set time validation for specific column in Excel using Aspose.Cells .NET | apply between operator for time values in Aspose.Cells validation | custom input and error messages for data validation in Aspose.Cells workbook
// Tags: Aspose.Cells time validation between operator | Excel column T validation Aspose.Cells | C# set time range validation Aspose.Cells | custom validation messages Aspose.Cells | apply validation rows 1-1000 Aspose.Cells

using System;
using Aspose.Cells;

// Creates a new workbook, defines a CellArea for column T rows 1‑1000, adds a Validation of type Time with the Between operator and bounds 09:00:00‑17:00:00, sets custom input and error titles/messages, and saves the file as TimeValidation.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define the validation area for column T (column index 19)
        // Here we apply it to rows 1 through 1000 (zero‑based indices 0‑999)
        CellArea area = CellArea.CreateCellArea(0, 19, 999, 19);

        // Add a validation object for the defined area
        ValidationCollection validations = sheet.Validations;
        int validationIndex = validations.Add(area);
        Validation validation = validations[validationIndex];

        // Configure the validation to allow only time values between 09:00 and 17:00
        validation.Type = ValidationType.Time;               // Time validation
        validation.Operator = OperatorType.Between;          // Between operator
        validation.Formula1 = "09:00:00";                    // Lower bound
        validation.Formula2 = "17:00:00";                    // Upper bound

        // Optional user‑friendly messages
        validation.InputTitle = "Time Entry";
        validation.InputMessage = "Enter a time between 09:00 and 17:00.";
        validation.ErrorTitle = "Invalid Time";
        validation.ErrorMessage = "The time must be between 09:00 and 17:00.";
        validation.ShowInput = true;
        validation.ShowError = true;
        validation.AlertStyle = ValidationAlertType.Stop;    // Stop alert on error

        // Save the workbook with the validation applied
        workbook.Save("TimeValidation.xlsx");
    }
}
