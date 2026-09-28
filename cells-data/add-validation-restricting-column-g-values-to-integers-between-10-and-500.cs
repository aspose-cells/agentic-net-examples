// Title: Apply integer validation (10‑500) to column G rows 0‑999 using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that adds a whole‑number validation to column G (rows 0‑999) restricting values to 10‑500 and includes custom input and error messages. | Generate a C# example that creates a validation rule for column G, sets the operator to Between, defines Formula1 = 10 and Formula2 = 500, and saves the workbook as ColumnG_Validation.xlsx.
// Common Searches: aspnet aspose.cells how to restrict column G values to integers between 10 and 500 | c# add data validation for a specific column in Excel using Aspose.Cells | set whole number validation range for column G rows 0 to 999 Aspose.Cells | aspose.cells validation with custom input and error messages example | c# create Excel file with integer range validation using Aspose.Cells library
// Tags: Aspose.Cells integer range validation C# | Excel column G whole number validation | data validation between operator Aspose.Cells | custom input error messages Aspose.Cells | validate rows 0-999 column G Aspose.Cells

using System;
using Aspose.Cells;

// Demonstrates how to use Aspose.Cells for .NET to add a whole‑number validation to column G (rows 0‑999), enforcing integer values between 10 and 500, with custom input and error prompts, and saves the result as an .xlsx workbook.
class ColumnGIntegerValidation
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the validation area for column G (zero‑based column index 6)
        // Here we apply the validation to rows 0‑999 (adjust as needed)
        CellArea validationArea = CellArea.CreateCellArea(0, 6, 999, 6);

        // Add a new validation to the worksheet for the defined area
        int validationIndex = worksheet.Validations.Add(validationArea);
        Validation validation = worksheet.Validations[validationIndex];

        // Configure the validation to allow whole numbers between 10 and 500
        validation.Type = ValidationType.WholeNumber;
        validation.Operator = OperatorType.Between;
        validation.Formula1 = "10";
        validation.Formula2 = "500";

        // Optional: set user‑friendly messages
        validation.InputTitle = "Enter Integer";
        validation.InputMessage = "Please enter an integer between 10 and 500.";
        validation.ErrorTitle = "Invalid Input";
        validation.ErrorMessage = "The value must be an integer between 10 and 500.";
        validation.ShowInput = true;
        validation.ShowError = true;

        // Save the workbook
        workbook.Save("ColumnG_Validation.xlsx");
    }
}
