using System;
using Aspose.Cells;

namespace AsposeCellsValidationExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the validation area for column V (index 21) from row 0 to row 999
            // Adjust the end row as needed for your data range
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 999,
                StartColumn = 21, // Column V
                EndColumn = 21
            };

            // Add a new validation to the worksheet for the defined area
            Validation validation = sheet.Validations[sheet.Validations.Add(area)];

            // Use AnyValue type and set IgnoreBlank to false to disallow blank entries
            validation.Type = ValidationType.AnyValue;
            validation.IgnoreBlank = false;

            // Configure the error alert that will be shown when a blank is entered
            validation.AlertStyle = ValidationAlertType.Stop; // Show a stop alert
            validation.ErrorTitle = "Invalid Entry";
            validation.ErrorMessage = "Blank values are not allowed in column V.";
            validation.ShowError = true;   // Ensure the error message is displayed
            validation.ShowInput = false;  // No input message needed

            // Save the workbook to a file
            workbook.Save("ColumnV_NoBlank_Validation.xlsx");
        }
    }
}