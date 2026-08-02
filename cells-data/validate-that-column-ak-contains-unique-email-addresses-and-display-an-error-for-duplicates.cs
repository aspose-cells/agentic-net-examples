// Title: Aspose.Cells C# – Enforce Unique Email Addresses in Column AK via Data Validation
// Description: Creates or loads a workbook, defines the data range for column AK (starting at row 2), adds a custom validation using the formula COUNTIF($AK:$AK,AK1)=1 to guarantee each email appears only once, configures a stop‑alert with a clear error message, and saves the file. Ideal for .NET developers needing Excel‑level uniqueness checks.
// Keywords: Aspose.Cells unique email validation C# | Excel column AK duplicate check | custom data validation COUNTIF Aspose | prevent duplicate emails .NET | Excel uniqueness rule Aspose.Cells | C# workbook data validation | email address uniqueness Excel
// Common Searches: Aspose.Cells enforce unique values in a column | C# data validation duplicate email Excel | COUNTIF custom validation Aspose.Cells | how to stop duplicate emails in Excel using .NET | set error message for duplicate cells Aspose
// Developer Intent: Add a custom data‑validation rule that blocks duplicate email entries in column AK and shows a stop‑alert when a duplicate is entered.
// Use Cases: Ensure a mailing‑list workbook contains only distinct customer emails. | Validate employee email IDs in an HR spreadsheet before distribution. | Maintain a supplier contact sheet with unique email addresses.
// AI Prompts: Generate C# code with Aspose.Cells that adds a custom validation to column AK to detect duplicate emails and display a stop alert. | Explain how the COUNTIF($AK:$AK,AK1)=1 formula works inside Aspose.Cells validation for uniqueness. | Show how to modify the validation range to start at row 5 instead of row 2.

using System;
using Aspose.Cells;

namespace AsposeCellsEmailValidation
{
    // Creates or loads a workbook, defines the data range for column AK (starting at row 2), adds a custom validation using the formula COUNTIF($AK:$AK,AK1)=1 to guarantee each email appears only once, configures a stop‑alert with a clear error message, and saves the file. Ideal for .NET developers needing Excel‑level uniqueness checks.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // lifecycle: create
            Worksheet worksheet = workbook.Worksheets[0];

            // Example data population (remove if loading an existing file)
            worksheet.Cells["AK2"].PutValue("alice@example.com");
            worksheet.Cells["AK3"].PutValue("bob@example.com");
            worksheet.Cells["AK4"].PutValue("alice@example.com"); // duplicate for demo

            // Determine the last row that contains data
            int lastDataRow = worksheet.Cells.MaxDataRow;
            if (lastDataRow < 1) lastDataRow = 1; // ensure at least one data row

            // Column AK is the 37th column (0‑based index 36)
            const int emailColumnIndex = 36;

            // Define the validation range (from row 2 to the last data row in column AK)
            CellArea emailRange = new CellArea
            {
                StartRow = 1,                 // row 2 (0‑based)
                EndRow = lastDataRow,
                StartColumn = emailColumnIndex,
                EndColumn = emailColumnIndex
            };

            // Add a custom data validation to the range
            int validationIdx = worksheet.Validations.Add(emailRange);
            Validation emailValidation = worksheet.Validations[validationIdx];

            // Set validation to ensure each email appears only once in column AK
            emailValidation.Type = ValidationType.Custom;
            // Formula checks that the count of the current cell's value in the whole column equals 1
            emailValidation.Formula1 = "COUNTIF($AK:$AK,AK1)=1";

            // Configure the error message that will be shown for duplicates
            emailValidation.ShowError = true;
            emailValidation.AlertStyle = ValidationAlertType.Stop;
            emailValidation.ErrorTitle = "Duplicate Email";
            emailValidation.ErrorMessage = "Email address must be unique in column AK.";

            // Save the workbook (lifecycle: save)
            workbook.Save("EmailValidationDemo.xlsx");
        }
    }
}
