// Title: Add numeric data validation (0‑500) to column K in an Aspose.Cells workbook (C#)
// Description: Creates a new workbook, defines a validation area for column K (rows 1‑1001), and applies a WholeNumber Between rule that restricts entries to values from 0 to 500, with custom input and error messages, then saves the file.
// Keywords: Aspose.Cells C# data validation | numeric validation column K | Excel whole number range 0 to 500 | Aspose.Cells validation operator Between | custom input error messages Aspose.Cells | C# create workbook with validation | Aspose.Cells .NET example
// Common Searches: How to limit values in Excel column K to 0‑500 using Aspose.Cells C# | Aspose.Cells add whole number validation between 0 and 500 | C# data validation for a specific column in Aspose.Cells workbook | Set numeric range validation in Aspose.Cells for .NET | Excel column validation example Aspose.Cells
// Developer Intent: Add a validation rule that permits only whole numbers from 0 to 500 in column K of a generated workbook.
// Use Cases: Automatically enforce price limits (0‑500) when generating invoice spreadsheets. | Validate sensor reading values stored in column K during data import. | Provide end‑users with clear prompts and error alerts for numeric entry in reports.
// AI Prompts: Write C# code using Aspose.Cells to apply a 0‑500 whole‑number validation to column K with custom messages. | Show how to adjust the validation range dynamically based on the last populated row in column K. | Explain how to export the workbook after adding data validation for column K in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsValidationExample
{
    // Creates a new workbook, defines a validation area for column K (rows 1‑1001), and applies a WholeNumber Between rule that restricts entries to values from 0 to 500, with custom input and error messages, then saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the validation area for column K (zero‑based column index 10)
            // Here we apply the validation to rows 0 through 1000; adjust as needed.
            CellArea validationArea = CellArea.CreateCellArea(0, 10, 1000, 10);

            // Add a new validation to the worksheet's validation collection
            int validationIndex = worksheet.Validations.Add(validationArea);
            Validation validation = worksheet.Validations[validationIndex];

            // Configure the validation: whole numbers between 0 and 500
            validation.Type = ValidationType.WholeNumber;
            validation.Operator = OperatorType.Between;
            validation.Formula1 = "0";
            validation.Formula2 = "500";

            // Optional: set messages and behavior
            validation.InputTitle = "Enter a number";
            validation.InputMessage = "Please enter a whole number between 0 and 500.";
            validation.ErrorTitle = "Invalid input";
            validation.ErrorMessage = "The value must be a whole number between 0 and 500.";
            validation.ShowInput = true;
            validation.ShowError = true;
            validation.AlertStyle = ValidationAlertType.Stop;
            validation.IgnoreBlank = true;
            validation.InCellDropDown = false;

            // Save the workbook
            workbook.Save("ColumnK_Validation.xlsx");
        }
    }
}
