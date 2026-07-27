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
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the allowed values for column S
            string[] allowedValues = new string[] { "Apple", "Banana", "Cherry", "Date" };

            // Create a cell area that covers column S (index 18) from row 0 to row 99
            // Adjust the end row as needed for your data range
            CellArea area = CellArea.CreateCellArea(0, 18, 99, 18);

            // Add a new validation to the worksheet for the defined area
            int validationIndex = worksheet.Validations.Add(area);
            Validation validation = worksheet.Validations[validationIndex];

            // Configure the validation as a list with the predefined values
            validation.Type = ValidationType.List;
            // Use comma‑separated string for the list values
            validation.Formula1 = string.Join(",", allowedValues);
            // Show the drop‑down list in the cell
            validation.InCellDropDown = true;

            // Optional: set an error message that appears when an invalid value is entered
            validation.AlertStyle = ValidationAlertType.Stop;
            validation.ErrorTitle = "Invalid Entry";
            validation.ErrorMessage = "Please select a value from the list.";

            // Save the workbook to a file
            workbook.Save("ColumnS_Validation.xlsx", SaveFormat.Xlsx);
        }
    }
}