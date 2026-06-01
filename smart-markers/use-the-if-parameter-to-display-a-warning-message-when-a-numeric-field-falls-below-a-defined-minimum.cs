using System;
using Aspose.Cells;

namespace AsposeCellsConditionalWarningDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range where the validation will be applied (e.g., cells A1:A10)
            CellArea validationArea = CellArea.CreateCellArea("A1", "A10");

            // Add a new validation rule to the worksheet
            Validation validation = sheet.Validations[sheet.Validations.Add()];
            validation.AddArea(validationArea);

            // Configure the validation to check for whole numbers greater than or equal to a minimum value
            validation.Type = ValidationType.WholeNumber;          // Validate whole numbers
            validation.Operator = OperatorType.LessThan;           // Trigger when value is less than the minimum
            validation.Formula1 = "10";                            // Minimum allowed value (e.g., 10)

            // Set the alert style to a warning (instead of stop)
            validation.AlertStyle = ValidationAlertType.Warning;   // Show a warning dialog
            validation.ShowError = true;                           // Ensure the warning is displayed
            validation.ErrorTitle = "Low Value Warning";
            validation.ErrorMessage = "The entered number is below the allowed minimum of 10.";

            // Optional: show input message when the cell is selected
            validation.ShowInput = true;
            validation.InputTitle = "Enter Value";
            validation.InputMessage = "Please enter a number greater than or equal to 10.";

            // Save the workbook to a file
            workbook.Save("ConditionalWarningDemo.xlsx");
        }
    }
}