// Title: C# – Apply 09:00‑17:00 Time Validation to Column T with Aspose.Cells
// Description: Shows how to create a workbook, define the whole column T (index 19) as a CellArea, add a Validation object, set its type to Time with a Between operator, and restrict entries to 09:00‑17:00. Includes optional input/error messages and saves the file as TimeValidation.xlsx.
// Keywords: Aspose.Cells | C# time validation | Excel column T validation | business hours data validation | ValidationType.Time | 09:00 to 17:00 Excel | data validation .NET | restrict Excel time entry
// Common Searches: Aspose.Cells set time validation column T | C# restrict Excel column to 09:00‑17:00 | add time range validation in .NET workbook | business hours data validation Aspose.Cells | create time validation for entire column C#
// Developer Intent: Create a validation rule that permits only times between 09:00 and 17:00 in column T.
// Use Cases: Ensure employee clock‑in times fall within standard work hours on a timesheet. | Limit appointment start times to office hours in a scheduling worksheet. | Prevent out‑of‑office timestamps in a project log by locking column T to business hours.
// AI Prompts: Generate Aspose.Cells C# code to enforce a 08:30‑18:30 time range in column B. | Show how to customize input and error messages for a time validation rule in Aspose.Cells. | Explain how to apply the same 09:00‑17:00 validation to multiple non‑adjacent columns using Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to create a workbook, define the whole column T (index 19) as a CellArea, add a Validation object, set its type to Time with a Between operator, and restrict entries to 09:00‑17:00. Includes optional input/error messages and saves the file as TimeValidation.xlsx.
class TimeValidationExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the validation area: entire column T (index 19)
            // Rows 0 to 1,048,575 cover the full Excel sheet
            CellArea timeArea = CellArea.CreateCellArea(0, 19, 1048575, 19);

            // Add a validation to the defined area
            int validationIdx = sheet.Validations.Add(timeArea);
            Validation timeValidation = sheet.Validations[validationIdx];

            // Set validation type to Time and require a value between 09:00 and 17:00
            timeValidation.Type = ValidationType.Time;
            timeValidation.Operator = OperatorType.Between;
            timeValidation.Formula1 = "09:00";
            timeValidation.Formula2 = "17:00";

            // Optional user messages
            timeValidation.InputMessage = "Enter a time between 09:00 and 17:00.";
            timeValidation.ErrorMessage = "Invalid time. Must be between 09:00 and 17:00.";
            timeValidation.ShowInput = true;
            timeValidation.ShowError = true;
            timeValidation.AlertStyle = ValidationAlertType.Stop;

            // Save the workbook
            workbook.Save("TimeValidation.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
