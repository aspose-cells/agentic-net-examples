// Title: C# – Add email validation with input prompt using Aspose.Cells for .NET
// Description: Creates a new workbook, defines cell A1, applies a custom validation formula that checks for "@" and a period, displays an input message asking for a valid email address, shows a stop‑style error alert for invalid entries, and saves the file as EmailValidationDemo.xlsx.
// Keywords: Aspose.Cells | C# | .NET | email validation | custom data validation | input message | error alert | Excel formula | worksheet validation | CellArea | ShowInput | ShowError | ValidationAlertType.Stop
// Common Searches: Aspose.Cells email validation with input message C# | How to show input prompt for data validation in Aspose.Cells | Custom validation formula for email address in .NET workbook | Set error alert for invalid email using Aspose.Cells | Add data validation to a single cell with Aspose.Cells
// Developer Intent: Generate a .NET workbook that validates an email address in cell A1, shows an input prompt, and displays an error alert for incorrect formats.
// Use Cases: Guide users to enter correctly formatted email addresses in a generated Excel report. | Prevent malformed email entries in a data‑entry template exported from a C# application. | Provide immediate feedback with a stop‑style alert when a user types an invalid email.
// AI Prompts: Write Aspose.Cells code to add phone‑number validation with an input message and error alert in C#. | Modify the email validation formula to require a domain suffix of at least two characters. | Apply the same email validation with input prompt to a range such as A1:A10 using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsEmailValidationDemo
{
    // Creates a new workbook, defines cell A1, applies a custom validation formula that checks for "@" and a period, displays an input message asking for a valid email address, shows a stop‑style error alert for invalid entries, and saves the file as EmailValidationDemo.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the cell (A1) where the email validation will be applied
            CellArea emailArea = new CellArea
            {
                StartRow = 0,    // Row 1 (zero‑based index)
                EndRow = 0,
                StartColumn = 0, // Column A (zero‑based index)
                EndColumn = 0
            };

            // Add a new validation to the worksheet and obtain the Validation object
            int validationIndex = worksheet.Validations.Add(emailArea);
            Validation emailValidation = worksheet.Validations[validationIndex];

            // Set validation type to Custom and provide a formula that checks for a basic email pattern
            emailValidation.Type = ValidationType.Custom;
            // The formula uses Excel's ISNUMBER and SEARCH functions to ensure the presence of "@" and "."
            emailValidation.SetFormula1("=AND(ISNUMBER(SEARCH(\"@\",A1)), ISNUMBER(SEARCH(\".\",A1)))", false, false);

            // Configure the input message that appears when the cell is selected
            emailValidation.ShowInput = true;
            emailValidation.InputTitle = "Email Address";
            emailValidation.InputMessage = "Please enter a valid email address (e.g., user@example.com)";

            // Configure the error alert that appears when an invalid value is entered
            emailValidation.ShowError = true;
            emailValidation.ErrorTitle = "Invalid Email";
            emailValidation.ErrorMessage = "The value must be a valid email address.";
            emailValidation.AlertStyle = ValidationAlertType.Stop;

            // Save the workbook
            workbook.Save("EmailValidationDemo.xlsx");
        }
    }
}
