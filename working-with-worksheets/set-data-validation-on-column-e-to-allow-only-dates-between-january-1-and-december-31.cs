// Title: Add date range validation to column E (E1:E1000) with Aspose.Cells for .NET
// Description: Shows how to create a Workbook, define a CellArea for rows 1‑1000 of column E, add a Validation object, set its type to Date with the Between operator, specify start (1/1/2023) and end (12/31/2023) dates, configure optional input and error messages, and save the file as ColumnEDateValidation.xlsx.
// Keywords: Aspose.Cells C# date validation | Excel column E data validation .NET | set date range Aspose.Cells | ValidationType.Date Between operator | custom input error messages Aspose.Cells | C# Excel validation example | apply validation to column E
// Common Searches: Aspose.Cells set date validation for column E | C# restrict Excel column to dates between Jan 1 and Dec 31 | How to add date range validation in Aspose.Cells .NET | Excel data validation column E Aspose.Cells example | Custom error message for date validation Aspose.Cells
// Developer Intent: Add a data‑validation rule that permits only dates from January 1 to December 31 in column E of an Excel worksheet.
// Use Cases: Ensure dates entered in a financial template fall within the fiscal year. | Prevent out‑of‑range dates when users upload scheduling data. | Create a guided entry form with clear prompts and error alerts for date fields.
// AI Prompts: Generate C# code using Aspose.Cells to apply a date‑range validation to column F for 2024, including custom input and error messages. | Show how to modify the validation to allow only weekdays between two dates in Aspose.Cells. | Provide an example of removing an existing validation from a column and replacing it with a new date‑range rule.

using System;
using Aspose.Cells;

namespace AsposeCellsValidationExample
{
    // Shows how to create a Workbook, define a CellArea for rows 1‑1000 of column E, add a Validation object, set its type to Date with the Between operator, specify start (1/1/2023) and end (12/31/2023) dates, configure optional input and error messages, and save the file as ColumnEDateValidation.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the validation area for column E (zero‑based column index 4)
            // Here we apply the validation to rows 0 through 999 (E1:E1000)
            CellArea validationArea = CellArea.CreateCellArea(0, 4, 999, 4);

            // Add a new validation to the worksheet
            int validationIndex = worksheet.Validations.Add(validationArea);
            Validation validation = worksheet.Validations[validationIndex];

            // Configure the validation to allow only dates between Jan 1 and Dec 31
            validation.Type = ValidationType.Date;                     // Date validation
            validation.Operator = OperatorType.Between;                // Between two dates
            validation.Formula1 = "1/1/2023";                           // Start date (Jan 1)
            validation.Formula2 = "12/31/2023";                         // End date (Dec 31)

            // Optional: user-friendly messages
            validation.InputTitle = "Date Required";
            validation.InputMessage = "Enter a date between Jan 1 and Dec 31.";
            validation.ErrorTitle = "Invalid Date";
            validation.ErrorMessage = "The date must be within the year 2023.";
            validation.ShowInput = true;
            validation.ShowError = true;

            // Save the workbook
            workbook.Save("ColumnEDateValidation.xlsx");
        }
    }
}
